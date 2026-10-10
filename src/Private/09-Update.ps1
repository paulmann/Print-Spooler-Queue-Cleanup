# ============================================================================
#  Update subsystem (finding A8)
# ----------------------------------------------------------------------------
#  Design goals, in priority order:
#    1. NEVER brick a working tool. The running script is never overwritten in
#       place while it executes; updates are STAGED into .update\ and swapped in
#       by the batch launcher (or by the next PowerShell run) before execution.
#    2. NEVER install unverified bytes. Assets are checked against the SHA-256
#       published in the release's checksums.txt. When no checksum is published
#       the operator must pass -SkipChecksumVerification explicitly.
#    3. Work behind rate limits and proxies. Three discovery strategies:
#         /releases/latest  ->  /tags  ->  raw branch file version stamp.
#       A GitHub token lifts the anonymous 60 req/h ceiling to 5000.
#    4. Be quiet by default. -AutoUpdate performs at most one check per cooldown
#       window and never blocks a cleanup for more than a couple of seconds.
# ============================================================================

function New-WebRequestCompat {
    <#
    .SYNOPSIS
        Invoke-WebRequest / Invoke-RestMethod wrapper that behaves identically on
        Windows PowerShell 5.1 and PowerShell 7.6.
    .DESCRIPTION
        5.1 needs -UseBasicParsing (the IE-engine dependency) and has no
        -SkipHttpErrorCheck; 7.x has both but also enables response caching that
        can hide a freshly published release. Normalising here keeps the callers
        readable. TLS 1.2 is forced on 5.1 because GitHub rejects TLS 1.0/1.1.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Uri,
        [AllowNull()][System.Collections.IDictionary]$Headers,
        [int]$TimeoutSec = $Script:UpdateTimeoutSeconds,
        [ValidateSet('Get', 'Head')][string]$Method = 'Get',
        [AllowNull()][AllowEmptyString()][string]$OutFile,
        [switch]$Raw,
        [switch]$AsJson
    )

    if (Test-RunningOnWindows) {
        try {
            if ($PSVersionTable.PSVersion.Major -le 5) {
                [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            }
        }
        catch {
            # An OS that cannot set the protocol will still negotiate TLS 1.2 by default.
        }
    }

    $allHeaders = @{
        'User-Agent'    = ($Script:UpdateUserAgent -f $Script:ScriptVersion, $Script:ScriptRepoUrl)
        'Accept'        = 'application/vnd.github+json'
        'Cache-Control' = 'no-cache'
    }
    if ($Headers) { foreach ($k in @($Headers.Keys)) { $allHeaders[[string]$k] = $Headers[$k] } }

    $params = @{
        Uri         = $Uri
        Method      = $Method
        Headers     = $allHeaders
        TimeoutSec  = $TimeoutSec
        ErrorAction = 'Stop'
    }
    if ($PSVersionTable.PSVersion.Major -le 5) { $params['UseBasicParsing'] = $true }
    if (-not [string]::IsNullOrWhiteSpace($OutFile)) { $params['OutFile'] = $OutFile }

    if ($AsJson) { return Invoke-RestMethod @params }

    $resp = Invoke-WebRequest @params
    if (-not [string]::IsNullOrWhiteSpace($OutFile)) { return $true }
    if ($resp.Content -is [byte[]]) { return [System.Text.Encoding]::UTF8.GetString($resp.Content) }
    return [string]$resp.Content
}

function Get-UpdateAuthenticationHeader {
    <#
    .SYNOPSIS
        Builds the Authorization header for a GitHub token, if one is supplied.
    .NOTES
        Tokens are read from -GitHubToken, PSQ_GITHUB_TOKEN or the config file and
        are never written to the log file, the report, or the diagnostics bundle.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param([AllowNull()][AllowEmptyString()][string]$Token)

    $h = @{}
    if ([string]::IsNullOrWhiteSpace($Token)) { return $h }
    $t = $Token.Trim()
    if ($t -match '^(ghp|gho|ghu|ghs|ghr|github-pat)_') { $h['Authorization'] = "Bearer $t" }
    else { $h['Authorization'] = "token $t" }
    $h['X-GitHub-Api-Version'] = '2022-11-28'
    return $h
}

function Get-VersionFromRemoteScript {
    <#
    .SYNOPSIS
        Strategy 3: parse the version stamp out of raw script text.
    .DESCRIPTION
        Works on repositories that publish no releases or tags at all - the state
        this repository was in before v4. Uses the AST rather than a regex over the
        whole file, so a version-looking string in a comment cannot fool it.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][AllowNull()][AllowEmptyString()][string]$Content,
        [string[]]$VariableNames = @('ScriptVersion', 'ModuleVersion')
    )

    if ([string]::IsNullOrWhiteSpace($Content)) { return $null }

    try {
        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($Content, [ref]$tokens, [ref]$errors)
        if ($ast) {
            foreach ($name in $VariableNames) {
                $escaped = [regex]::Escape($name)
                $assign = @($ast.FindAll({
                            param($node)
                            ($node -is [System.Management.Automation.Language.AssignmentStatementAst]) -and
                            ($node.Left.Extent.Text -match $escaped)
                        }, $true) | Select-Object -First 1)
                if ($assign.Count -gt 0 -and $assign[0]) {
                    $text = $assign[0].Right.Extent.Text.Trim().Trim('"', "'")
                    if ($text -match '^\d+(\.\d+){0,3}') { return $Matches[0] }
                }
            }
        }
    }
    catch {
        # Fall through to the regex below; a parse failure is not fatal here.
    }

    # Last-ditch pattern for scripts that only carry the version in .NOTES.
    if ($Content -match '(?m)^\s*(?:Script\s*)?Version\s*[:=]\s*v?(\d+\.\d+(?:\.\d+){0,2})') { return $Matches[1] }
    return $null
}

function ConvertFrom-ChecksumText {
    <#
    .SYNOPSIS
        Parses a sha256sum-style file into @{ 'file.ext' = 'HEX' }.
    .DESCRIPTION
        Accepts the GNU coreutils form ("<hex>  <name>"), the reversed
        "<name>: <hex>" form, the "*binary" marker, and comment lines.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param([Parameter(Mandatory)][AllowNull()][AllowEmptyString()][string]$Text)

    $map = @{}
    if ([string]::IsNullOrWhiteSpace($Text)) { return $map }
    foreach ($line in ($Text -split "\r?\n")) {
        $l = $line.Trim()
        if (-not $l -or $l.StartsWith('#')) { continue }
        $l = $l -replace '\*', ' '
        if ($l -match '^(?<hash>[0-9A-Fa-f]{32,128})\s+(?<name>.+)$') {
            $map[$Matches['name'].Trim()] = $Matches['hash']
        }
        elseif ($l -match '^(?<name>[^\s:]+):\s*(?<hash>[0-9A-Fa-f]{32,128})\s*$') {
            $map[$Matches['name'].Trim()] = $Matches['hash']
        }
    }
    return $map
}

function ConvertTo-ChecksumText {
    <#
    .SYNOPSIS
        Produces a GNU-coreutils-compatible checksums.txt for a release folder.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][AllowNull()][AllowEmptyCollection()][string[]]$FilePaths)

    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($p in @($FilePaths)) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        if (-not (Test-Path -LiteralPath $p -PathType Leaf)) { continue }
        $h = Get-FileHashSafe -Path $p -Algorithm 'SHA256'
        if ($h) { $lines.Add(("{0}  {1}" -f $h.ToLowerInvariant(), [IO.Path]::GetFileName($p))) }
    }
    return (($lines -join "`n") + "`n")
}

function Get-GitHubLatestRelease {
    <#
    .SYNOPSIS
        Strategy 1: query /releases/latest (or /releases for the other channels).
    .OUTPUTS
        PSCustomObject { TagName, Version, Name, PublishedUtc, Prerelease, Assets, Body, Source, Error }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Owner,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Repo,
        [ValidateSet('latest', 'prerelease', 'stable')][string]$Channel = 'latest',
        [AllowNull()][AllowEmptyString()][string]$Token
    )

    $headers = Get-UpdateAuthenticationHeader -Token $Token
    $result = [pscustomobject]@{
        TagName      = $null
        Version      = $null
        Name         = $null
        PublishedUtc = $null
        Prerelease   = $false
        Assets       = @()
        Body         = $null
        Source       = 'releases'
        Error        = $null
    }

    $mapAssets = {
        param($list)
        @($list | ForEach-Object {
                [pscustomobject]@{
                    Name        = [string]$_.name
                    Url         = [string]$_.browser_download_url
                    Size        = [long]$_.size
                    ContentType = [string]$_.content_type
                }
            })
    }

    try {
        if ($Channel -eq 'latest') {
            $uri = "{0}/{1}/{2}/releases/latest" -f $Script:GitHubApiBase, $Owner, $Repo
            $rel = New-WebRequestCompat -Uri $uri -Headers $headers -AsJson
            $result.TagName = [string]$rel.tag_name
            $result.Name = [string]$rel.name
            $result.PublishedUtc = [string]$rel.published_at
            $result.Prerelease = [bool]$rel.prerelease
            $result.Body = [string]$rel.body
            $result.Assets = (& $mapAssets $rel.assets)
        }
        else {
            $uri = "{0}/{1}/{2}/releases?per_page=30" -f $Script:GitHubApiBase, $Owner, $Repo
            $rels = @(New-WebRequestCompat -Uri $uri -Headers $headers -AsJson)
            $pick = $null
            foreach ($r in $rels) {
                if ([bool]$r.draft) { continue }
                if (($Channel -eq 'stable') -and [bool]$r.prerelease) { continue }
                $pick = $r
                break
            }
            if (-not $pick) {
                $result.Error = 'No matching release was found.'
                return $result
            }
            $result.TagName = [string]$pick.tag_name
            $result.Name = [string]$pick.name
            $result.PublishedUtc = [string]$pick.published_at
            $result.Prerelease = [bool]$pick.prerelease
            $result.Body = [string]$pick.body
            $result.Assets = (& $mapAssets $pick.assets)
        }
        $sv = ConvertTo-SemVer -Value $result.TagName
        if ($sv) { $result.Version = $sv.Normalized }
    }
    catch {
        $result.Error = $_.Exception.Message
        if ($_.Exception.Message -match '403|rate limit') {
            $result.Error = "GitHub API rate limit or forbidden (403). Set PSQ_GITHUB_TOKEN, or rely on the raw-file fallback. Detail: $($_.Exception.Message)"
        }
    }
    return $result
}

function Get-GitHubLatestTag {
    <#
    .SYNOPSIS
        Strategy 2: newest tag via /tags, for repositories with no formal releases.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Owner,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Repo,
        [AllowNull()][AllowEmptyString()][string]$Token,
        [switch]$IncludePrerelease
    )

    $headers = Get-UpdateAuthenticationHeader -Token $Token
    try {
        $uri = "{0}/{1}/{2}/tags?per_page=50" -f $Script:GitHubApiBase, $Owner, $Repo
        $tags = @(New-WebRequestCompat -Uri $uri -Headers $headers -AsJson)
        $best = $null
        foreach ($t in $tags) {
            $tagName = [string]$t.name
            if ([string]::IsNullOrWhiteSpace($tagName)) { continue }
            $sv = ConvertTo-SemVer -Value $tagName
            if (-not $sv) { continue }
            if ($sv.IsPrerelease -and -not $IncludePrerelease) { continue }
            if (-not $best -or ($sv.Version -gt $best.Version)) { $best = $sv }
        }
        if (-not $best) {
            return [pscustomobject]@{ TagName = $null; Version = $null; Prerelease = $false; Assets = @(); Source = 'tags'; Error = 'No usable tag was found.' }
        }
        return [pscustomobject]@{ TagName = $best.Raw; Version = $best.Normalized; Prerelease = $best.IsPrerelease; Assets = @(); Source = 'tags'; Error = $null }
    }
    catch {
        return [pscustomobject]@{ TagName = $null; Version = $null; Prerelease = $false; Assets = @(); Source = 'tags'; Error = $_.Exception.Message }
    }
}

function Get-RemoteVersionFromRawFile {
    <#
    .SYNOPSIS
        Strategy 3: fetch the raw script from the branch and read its version stamp.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Owner,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Repo,
        [string]$Branch = $Script:UpdateDefaultBranch,
        [string]$FileName = 'Clear-PrintSpoolerQueue.ps1'
    )

    $uri = "{0}/{1}/{2}/{3}/{4}" -f $Script:GitHubRawBase, $Owner, $Repo, $Branch, $FileName
    try {
        $content = New-WebRequestCompat -Uri $uri -Raw
        $ver = Get-VersionFromRemoteScript -Content $content
        if (-not $ver) {
            return [pscustomobject]@{ Version = $null; Source = 'raw'; Error = 'No version stamp was found in the remote script.'; RawUrl = $uri; Content = $content }
        }
        return [pscustomobject]@{ Version = $ver; Source = 'raw'; Error = $null; RawUrl = $uri; Content = $content }
    }
    catch {
        return [pscustomobject]@{ Version = $null; Source = 'raw'; Error = $_.Exception.Message; RawUrl = $uri; Content = $null }
    }
}

function Get-UpdateInfoCore {
    <#
    .SYNOPSIS
        Full discovery pipeline: releases -> tags -> raw branch file.
    .OUTPUTS
        PSCustomObject { CurrentVersion, LatestVersion, UpdateAvailable, Channel,
        Source, TagName, ReleaseName, PublishedUtc, Prerelease, ReleaseNotes,
        Assets, DownloadUrl, ExpectedSha256, ChecksumSource, RawUrl, Errors }
    #>
    [CmdletBinding()]
    param(
        [string]$Owner = $Script:ScriptRepoOwner,
        [string]$Repo = $Script:ScriptRepoName,
        [string]$Branch = $Script:UpdateDefaultBranch,
        [string]$CurrentVersion = $Script:ScriptVersion,
        [ValidateSet('latest', 'prerelease', 'stable')][string]$Channel = 'latest',
        [switch]$IncludePrerelease,
        [AllowNull()][AllowEmptyString()][string]$Token,
        [string]$PrimaryAssetName = 'Clear-PrintSpoolerQueue.ps1'
    )

    $errors = New-Object System.Collections.Generic.List[string]
    $info = [ordered]@{
        CurrentVersion  = $CurrentVersion
        LatestVersion   = $null
        UpdateAvailable = $false
        Channel         = $Channel
        Source          = $null
        TagName         = $null
        ReleaseName     = $null
        PublishedUtc    = $null
        Prerelease      = $false
        ReleaseNotes    = $null
        Assets          = @()
        DownloadUrl     = $null
        ExpectedSha256  = $null
        ChecksumSource  = $null
        RawUrl          = $null
        RepositoryUrl   = $Script:ScriptRepoUrl
        CheckedUtc      = [DateTime]::UtcNow.ToString('o')
        Errors          = @()
    }

    $headers = Get-UpdateAuthenticationHeader -Token $Token

    # ---- strategy 1: releases ------------------------------------------------
    $rel = Get-GitHubLatestRelease -Owner $Owner -Repo $Repo -Channel $Channel -Token $Token
    if ($rel.Version) {
        $info.LatestVersion = $rel.Version
        $info.TagName = $rel.TagName
        $info.ReleaseName = $rel.Name
        $info.PublishedUtc = $rel.PublishedUtc
        $info.Prerelease = $rel.Prerelease
        $info.ReleaseNotes = $rel.Body
        $info.Source = "releases/$Channel"
        $info.Assets = @($rel.Assets)

        $asset = @($rel.Assets | Where-Object { $_.Name -eq $PrimaryAssetName } | Select-Object -First 1)
        if ($asset.Count -eq 0) { $asset = @($rel.Assets | Where-Object { $_.Name -like '*.ps1' } | Select-Object -First 1) }
        if ($asset.Count -gt 0) { $info.DownloadUrl = $asset[0].Url }

        $zip = @($rel.Assets | Where-Object { $_.Name -like '*.zip' } | Select-Object -First 1)
        if (-not $info.DownloadUrl -and $zip.Count -gt 0) { $info.DownloadUrl = $zip[0].Url }

        $sums = @($rel.Assets | Where-Object { $_.Name -eq $Script:ChecksumAssetName } | Select-Object -First 1)
        if ($sums.Count -gt 0) {
            try {
                $txt = New-WebRequestCompat -Uri $sums[0].Url -Headers $headers -Raw
                $parsed = ConvertFrom-ChecksumText -Text $txt
                $target = $null
                if ($asset.Count -gt 0) { $target = $asset[0].Name } elseif ($zip.Count -gt 0) { $target = $zip[0].Name }
                if ($target -and $parsed.Contains($target)) {
                    $info.ExpectedSha256 = $parsed[$target]
                    $info.ChecksumSource = 'release-asset'
                }
            }
            catch {
                $errors.Add("checksums.txt could not be fetched: $($_.Exception.Message)")
            }
        }
    }
    elseif ($rel.Error) {
        $errors.Add("releases: $($rel.Error)")
    }

    # ---- strategy 2: tags ----------------------------------------------------
    if (-not $info.LatestVersion) {
        $tag = Get-GitHubLatestTag -Owner $Owner -Repo $Repo -Token $Token -IncludePrerelease:$IncludePrerelease
        if ($tag -and $tag.Version) {
            $info.LatestVersion = $tag.Version
            $info.TagName = $tag.TagName
            $info.Source = 'tags'
            $info.Prerelease = $tag.Prerelease
            $info.DownloadUrl = "{0}/{1}/{2}/refs/tags/{3}/{4}" -f $Script:GitHubRawBase, $Owner, $Repo, $tag.TagName, $PrimaryAssetName
            $sumUrl = "{0}/{1}/{2}/refs/tags/{3}/{4}" -f $Script:GitHubRawBase, $Owner, $Repo, $tag.TagName, $Script:ChecksumAssetName
            try {
                $txt = New-WebRequestCompat -Uri $sumUrl -Raw
                $parsed = ConvertFrom-ChecksumText -Text $txt
                if ($parsed.Contains($PrimaryAssetName)) {
                    $info.ExpectedSha256 = $parsed[$PrimaryAssetName]
                    $info.ChecksumSource = 'tag-file'
                }
            }
            catch {
                # A tag without checksums is still usable; the operator is told later.
            }
        }
        elseif ($tag -and $tag.Error) {
            $errors.Add("tags: $($tag.Error)")
        }
    }

    # ---- strategy 3: raw branch file ----------------------------------------
    if (-not $info.LatestVersion) {
        $raw = Get-RemoteVersionFromRawFile -Owner $Owner -Repo $Repo -Branch $Branch -FileName $PrimaryAssetName
        if ($raw.Version) {
            $info.LatestVersion = $raw.Version
            $info.Source = "raw/$Branch"
            $info.RawUrl = $raw.RawUrl
            $info.DownloadUrl = $raw.RawUrl
            $info.ReleaseNotes = 'No GitHub release metadata is available; the version was read from the branch source file.'
        }
        elseif ($raw.Error) {
            $errors.Add("raw: $($raw.Error)")
        }
    }

    if ($info.LatestVersion) {
        $info.UpdateAvailable = Test-NewerVersion -Current $CurrentVersion -Candidate $info.LatestVersion -IncludePrerelease:$IncludePrerelease
    }
    $info.Errors = $errors.ToArray()
    return [pscustomobject]$info
}

function Get-UpdateStatePath {
    <#
    .SYNOPSIS
        Where the update cooldown state file lives (first writable location).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()][AllowEmptyString()][string]$ScriptDirectory)

    $root = $null
    if (-not [string]::IsNullOrWhiteSpace($ScriptDirectory) -and (Test-DirectoryWritable -Path $ScriptDirectory)) { $root = $ScriptDirectory }
    elseif ($env:ProgramData -and (Test-DirectoryWritable -Path $env:ProgramData)) { $root = Join-Path $env:ProgramData $Script:ProductFamily }
    elseif ($env:LOCALAPPDATA) { $root = Join-Path $env:LOCALAPPDATA $Script:ProductFamily }
    else { $root = [IO.Path]::GetTempPath() }

    if (-not (Test-Path -LiteralPath $root)) {
        try { New-Item -ItemType Directory -Path $root -Force -ErrorAction Stop | Out-Null }
        catch { $root = [IO.Path]::GetTempPath() }
    }
    return (Join-Path $root $Script:UpdateStateFileName)
}

function Read-UpdateState {
    <#
    .SYNOPSIS
        Reads the update cooldown state file. Missing or corrupt = empty state.
    #>
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Path)

    $state = @{
        LastCheckUtc     = $null
        LastKnownVersion = $null
        LastCheckResult  = $null
        SuppressUntilUtc = $null
        InstallDirectory = $null
    }
    if ([string]::IsNullOrWhiteSpace($Path)) { return $state }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $state }

    try {
        $json = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        # The key list MUST be snapshotted. Enumerating $state.Keys while assigning
        # $state[$k] mutates the Hashtable's version counter and throws "Collection
        # was modified" mid-loop; combined with a catch-all that reports nothing,
        # this silently returned an EMPTY state - which disabled the update
        # cooldown entirely and made every run hit the GitHub API.
        foreach ($k in @($state.Keys)) {
            $prop = $json.PSObject.Properties[$k]
            if ($prop -and $null -ne $prop.Value) { $state[$k] = [string]$prop.Value }
        }
    }
    catch {
        # A corrupt state file must never break a run; it only costs one re-check.
        Write-LogDebug "The update state file '$Path' could not be read ($($_.Exception.Message)); treating it as empty." -Component 'Update'
    }
    return $state
}

function Write-UpdateState {
    <#
    .SYNOPSIS
        Persists the update cooldown state.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [AllowNull()][AllowEmptyString()][string]$Path,
        [Parameter(Mandatory)][AllowNull()][System.Collections.IDictionary]$State
    )
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    try {
        ($State | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop
        return $true
    }
    catch {
        Write-LogDebug "Could not persist the update state to '$Path': $($_.Exception.Message)" -Component 'Update'
        return $false
    }
}

function Test-UpdateCooldownActive {
    <#
    .SYNOPSIS
        True when the last check is still inside the cooldown window.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [AllowNull()][System.Collections.IDictionary]$State,
        [int]$CooldownHours = $Script:UpdateCooldownHours
    )
    if ($CooldownHours -le 0) { return $false }
    if (-not $State -or -not $State.Contains('LastCheckUtc') -or -not $State['LastCheckUtc']) { return $false }

    $last = [DateTime]::MinValue
    if (-not [DateTime]::TryParse([string]$State['LastCheckUtc'], [ref]$last)) { return $false }
    return (([DateTime]::UtcNow - $last.ToUniversalTime()).TotalHours -lt $CooldownHours)
}

function Invoke-PendingUpdateInternal {
    <#
    .SYNOPSIS
        Applies a staged update: .update\<file> -> <install>\<file>.
    .DESCRIPTION
        Called by the PowerShell entry point at start-up and by the batch launcher.
        Files are MOVED, not copied, so a half-applied update cannot be re-applied.
        Each staged file's SHA-256 is re-verified against the value recorded when
        it was staged, so a tampered .update directory cannot be promoted.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowNull()][AllowEmptyString()][string]$InstallDirectory,
        [switch]$Force
    )

    $out = [ordered]@{ Applied = $false; ToVersion = $null; Files = @(); Message = $null }
    if ([string]::IsNullOrWhiteSpace($InstallDirectory)) {
        $out.Message = 'No install directory.'
        return [pscustomobject]$out
    }

    $stageDir = Join-Path $InstallDirectory $Script:PendingUpdateDirName
    $marker = Join-Path $stageDir $Script:PendingUpdateMarker
    if (-not (Test-Path -LiteralPath $marker -PathType Leaf)) {
        $out.Message = 'No pending update.'
        return [pscustomobject]$out
    }

    try {
        $manifest = Get-Content -LiteralPath $marker -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
        if ($manifest.Applied -and -not $Force) {
            $out.Message = 'The pending update was already applied.'
            return [pscustomobject]$out
        }

        $appliedFiles = New-Object System.Collections.Generic.List[string]
        foreach ($f in @($manifest.Files)) {
            $src = Join-Path $stageDir $f.Name
            if (-not (Test-Path -LiteralPath $src -PathType Leaf)) { continue }
            if ($f.Sha256) {
                $actual = Get-FileHashSafe -Path $src -Algorithm 'SHA256'
                if (-not (Test-HashMatch -Expected $f.Sha256 -Actual $actual)) {
                    throw "The staged file '$($f.Name)' failed its integrity check; the pending update was discarded."
                }
            }
            $dst = Join-Path $InstallDirectory $f.Name
            try {
                Move-Item -LiteralPath $src -Destination $dst -Force -ErrorAction Stop
                $appliedFiles.Add($f.Name)
            }
            catch {
                # Locked because the script is running. Leave it staged for next start.
                $out.Message = "The file '$($f.Name)' is in use; the update stays staged and will be applied on the next start."
                return [pscustomobject]$out
            }
        }

        $manifest | Add-Member -NotePropertyName Applied -NotePropertyValue $true -Force
        $manifest | Add-Member -NotePropertyName AppliedUtc -NotePropertyValue ([DateTime]::UtcNow.ToString('o')) -Force
        ($manifest | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $marker -Encoding UTF8 -Force -ErrorAction SilentlyContinue

        $out.Applied = $true
        $out.ToVersion = $manifest.ToVersion
        $out.Files = $appliedFiles.ToArray()
        $out.Message = "Applied the pending update to v$($manifest.ToVersion) ($($appliedFiles.Count) file(s))."
        return [pscustomobject]$out
    }
    catch {
        $out.Message = $_.Exception.Message
        try { Remove-Item -LiteralPath $stageDir -Recurse -Force -ErrorAction SilentlyContinue }
        catch { /* the stage directory is inside our own install folder; leaving it is harmless */ }
        return [pscustomobject]$out
    }
}

function Update-SpoolerCleanupCore {
    <#
    .SYNOPSIS
        Downloads, verifies and stages (or installs) a newer version.
    .DESCRIPTION
        Staging model: new files land in <install>\.update\ together with
        pending-update.json. The batch launcher applies the swap on the next start;
        Clear-PrintSpoolerQueue.ps1 applies it at the top of its next run. That is
        what makes self-update safe - a process never overwrites the file it is
        executing from.
    .OUTPUTS
        PSCustomObject { Updated, Staged, FromVersion, ToVersion, DownloadedFiles,
        BackupPath, VerifyResult, Method, Message, Files }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowNull()]$UpdateInfo,
        [Parameter(Mandatory)][AllowEmptyString()][string]$InstallDirectory,
        [switch]$Force,
        [switch]$SkipChecksumVerification,
        [switch]$ApplyNow,
        [AllowNull()][AllowEmptyString()][string]$Token,
        [string[]]$FileNames = @('Clear-PrintSpoolerQueue.ps1', 'Clear-PrintSpoolerQueue.bat')
    )

    $result = [ordered]@{
        Updated         = $false
        Staged          = $false
        FromVersion     = $Script:ScriptVersion
        ToVersion       = $UpdateInfo.LatestVersion
        DownloadedFiles = @()
        BackupPath      = $null
        VerifyResult    = 'NotVerified'
        Method          = $null
        Message         = $null
        Files           = @()
    }

    if (-not $UpdateInfo.UpdateAvailable -and -not $Force) {
        $result.Message = "Already up to date (current $($Script:ScriptVersion), latest $($UpdateInfo.LatestVersion))."
        return [pscustomobject]$result
    }

    $stageDir = Join-Path $InstallDirectory $Script:PendingUpdateDirName
    $tempDir = $null
    try {
        if (-not (Test-DirectoryWritable -Path $InstallDirectory)) {
            throw "The install directory is not writable: $InstallDirectory. Re-run elevated, or download manually from $($UpdateInfo.RepositoryUrl)/releases."
        }
        if (-not (Test-Path -LiteralPath $stageDir)) {
            New-Item -ItemType Directory -Path $stageDir -Force -ErrorAction Stop | Out-Null
        }

        $tempDir = Join-Path ([IO.Path]::GetTempPath()) ("psq-update-{0}" -f [Guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $tempDir -Force -ErrorAction Stop | Out-Null

        $headers = Get-UpdateAuthenticationHeader -Token $Token
        $downloaded = New-Object System.Collections.Generic.List[string]

        # ---- source selection
        if ($UpdateInfo.Assets -and @($UpdateInfo.Assets).Count -gt 0) {
            foreach ($name in $FileNames) {
                $asset = @($UpdateInfo.Assets | Where-Object { $_.Name -eq $name } | Select-Object -First 1)
                if ($asset.Count -eq 0) { continue }
                $dest = Join-Path $tempDir $name
                Write-LogInfo "Downloading the release asset '$name' ($([Math]::Round($asset[0].Size / 1KB, 1)) KB)..." -Component 'Update'
                $null = New-WebRequestCompat -Uri $asset[0].Url -Headers $headers -OutFile $dest -TimeoutSec 120
                $downloaded.Add($dest)
            }
            $result.Method = 'release-assets'
        }
        elseif ($UpdateInfo.DownloadUrl) {
            $name = [IO.Path]::GetFileName(($UpdateInfo.DownloadUrl -split '\?')[0])
            if ([string]::IsNullOrWhiteSpace($name)) { $name = $FileNames[0] }
            $dest = Join-Path $tempDir $name
            Write-LogInfo "Downloading '$name' from $($UpdateInfo.DownloadUrl)..." -Component 'Update'
            $null = New-WebRequestCompat -Uri $UpdateInfo.DownloadUrl -Headers $headers -OutFile $dest -TimeoutSec 120
            $downloaded.Add($dest)
            $result.Method = 'direct-download'
        }
        else {
            throw 'No downloadable asset or URL was discovered for this release.'
        }

        if ($downloaded.Count -eq 0) {
            throw "None of the expected files ($($FileNames -join ', ')) were published with this release."
        }

        # ---- integrity verification
        if ($UpdateInfo.ExpectedSha256) {
            $primary = @($downloaded | Where-Object { [IO.Path]::GetFileName($_) -eq $FileNames[0] } | Select-Object -First 1)
            if ($primary.Count -eq 0) { $primary = @($downloaded[0]) }
            $actual = Get-FileHashSafe -Path $primary[0] -Algorithm 'SHA256'
            if (Test-HashMatch -Expected $UpdateInfo.ExpectedSha256 -Actual $actual) {
                $result.VerifyResult = 'ChecksumMatch'
                Write-LogSuccess "SHA-256 verified ($($UpdateInfo.ChecksumSource)): $actual" -Component 'Update'
            }
            else {
                $result.VerifyResult = 'ChecksumMismatch'
                throw ("Integrity check FAILED for '{0}'.`n  expected: {1}`n  actual  : {2}`nThe download was discarded and the installed copy is untouched. This usually means a tampered mirror or a broken release." -f `
                    [IO.Path]::GetFileName($primary[0]), $UpdateInfo.ExpectedSha256, $actual)
            }
        }
        elseif ($SkipChecksumVerification) {
            $result.VerifyResult = 'SkippedByOperator'
            Write-LogWarn 'No published checksum was found and -SkipChecksumVerification was supplied; this download is NOT integrity-verified.' -Component 'Update'
        }
        else {
            $result.VerifyResult = 'NoChecksumPublished'
            throw "This release publishes no $($Script:ChecksumAssetName), so the download cannot be verified. Either publish checksums with the release (tools/Build.ps1 generates them) or re-run with -SkipChecksumVerification."
        }

        # ---- sanity: the downloaded script must parse and carry a version
        $psCandidate = @($downloaded | Where-Object { $_ -like '*.ps1' } | Select-Object -First 1)
        if ($psCandidate.Count -gt 0) {
            $content = Get-Content -LiteralPath $psCandidate[0] -Raw -Encoding UTF8
            $remoteVersion = Get-VersionFromRemoteScript -Content $content
            $tokens = $null
            $perr = $null
            $null = [System.Management.Automation.Language.Parser]::ParseInput($content, [ref]$tokens, [ref]$perr)
            if ($perr -and $perr.Count -gt 0) {
                throw "The downloaded script has parse errors and was discarded: $($perr[0].Message)"
            }
            if ($remoteVersion -and ((Compare-ProductVersion -Reference $Script:ScriptVersion -Difference $remoteVersion) -ge 0)) {
                Write-LogWarn ("The downloaded script reports version {0}, which is not newer than the installed {1}. Continuing because the release metadata says {2}." -f `
                    $remoteVersion, $Script:ScriptVersion, $UpdateInfo.LatestVersion) -Component 'Update'
            }
        }

        # ---- back up the current files so a rollback is always possible
        $stamp = Get-SafeFileNameStamp
        $backupDir = Join-Path $InstallDirectory ("backup-{0}-{1}" -f $Script:ScriptVersion, $stamp)
        New-Item -ItemType Directory -Path $backupDir -Force -ErrorAction Stop | Out-Null
        foreach ($name in $FileNames) {
            $current = Join-Path $InstallDirectory $name
            if (Test-Path -LiteralPath $current -PathType Leaf) {
                Copy-Item -LiteralPath $current -Destination (Join-Path $backupDir $name) -Force -ErrorAction SilentlyContinue
            }
        }
        $result.BackupPath = $backupDir

        # ---- stage
        $stagedFiles = New-Object System.Collections.Generic.List[object]
        foreach ($f in $downloaded.ToArray()) {
            $name = [IO.Path]::GetFileName($f)
            $target = Join-Path $stageDir $name
            Copy-Item -LiteralPath $f -Destination $target -Force -ErrorAction Stop
            $stagedFiles.Add([pscustomobject]@{
                    Name   = $name
                    Size   = (Get-Item -LiteralPath $target -Force).Length
                    Sha256 = (Get-FileHashSafe -Path $target -Algorithm 'SHA256')
                })
        }

        $manifest = [ordered]@{
            CreatedUtc     = [DateTime]::UtcNow.ToString('o')
            FromVersion    = $Script:ScriptVersion
            ToVersion      = $UpdateInfo.LatestVersion
            TagName        = $UpdateInfo.TagName
            Source         = $UpdateInfo.Source
            VerifyResult   = $result.VerifyResult
            ExpectedSha256 = $UpdateInfo.ExpectedSha256
            BackupPath     = $backupDir
            Applied        = $false
            Files          = $stagedFiles.ToArray()
        }
        ($manifest | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath (Join-Path $stageDir $Script:PendingUpdateMarker) -Encoding UTF8 -ErrorAction Stop

        $result.Files = $stagedFiles.ToArray()
        $result.Staged = $true
        $result.Message = "v$($UpdateInfo.LatestVersion) is staged in '$stageDir' and will be applied automatically on the next start."

        if ($ApplyNow) {
            $applied = Invoke-PendingUpdateInternal -InstallDirectory $InstallDirectory -Force
            if ($applied.Applied) {
                $result.Updated = $true
                $result.Message = "Updated to v$($UpdateInfo.LatestVersion) in place. Restart the tool to run the new version."
            }
            else {
                $result.Message += " (In-place apply was not possible: $($applied.Message))"
            }
        }
    }
    catch {
        $result.Message = $_.Exception.Message
        if ($result.VerifyResult -eq 'NotVerified') { $result.VerifyResult = 'Failed' }
        Write-LogError "Update failed: $($_.Exception.Message)" -Component 'Update'
        throw
    }
    finally {
        if ($tempDir) { Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
    }
    return [pscustomobject]$result
}

function Invoke-AutoUpdateCheck {
    <#
    .SYNOPSIS
        Quiet, cooldown-limited check used at start-up (-AutoUpdate).
    .DESCRIPTION
        Returns $null when the check was skipped (cooldown, disabled, offline) so
        the caller can continue the cleanup without branching. A failed or slow
        check must never delay or abort a print-queue repair.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyString()][string]$InstallDirectory,
        [int]$CooldownHours = $Script:UpdateCooldownHours,
        [ValidateSet('latest', 'prerelease', 'stable')][string]$Channel = 'latest',
        [AllowNull()][AllowEmptyString()][string]$Token,
        [switch]$ApplyNow
    )

    $statePath = Get-UpdateStatePath -ScriptDirectory $InstallDirectory
    $state = Read-UpdateState -Path $statePath
    if (Test-UpdateCooldownActive -State $state -CooldownHours $CooldownHours) {
        Write-LogTrace "Update check skipped (cooldown active; last check $($state['LastCheckUtc']))." -Component 'Update'
        return $null
    }

    try {
        $info = Get-UpdateInfoCore -Channel $Channel -Token $Token
        $state['LastCheckUtc'] = [DateTime]::UtcNow.ToString('o')
        $state['LastKnownVersion'] = $info.LatestVersion
        $state['LastCheckResult'] = if ($info.UpdateAvailable) { 'UpdateAvailable' } else { 'UpToDate' }
        $state['InstallDirectory'] = $InstallDirectory
        $null = Write-UpdateState -Path $statePath -State $state

        if (-not $info.UpdateAvailable) {
            Write-LogDebug "Update check: v$($Script:ScriptVersion) is current (latest v$($info.LatestVersion) via $($info.Source))." -Component 'Update'
            return $info
        }

        Write-LogWarn ("A newer version is available: v{0} (this host runs v{1}). Source: {2}" -f $info.LatestVersion, $Script:ScriptVersion, $info.Source) -Component 'Update'
        if ($ApplyNow -and -not [string]::IsNullOrWhiteSpace($InstallDirectory)) {
            try {
                $upd = Update-SpoolerCleanupCore -UpdateInfo $info -InstallDirectory $InstallDirectory -ApplyNow -Token $Token
                Write-LogSuccess $upd.Message -Component 'Update'
            }
            catch {
                Write-LogWarn "The auto-update could not be applied: $($_.Exception.Message)" -Component 'Update'
            }
        }
        return $info
    }
    catch {
        Write-LogDebug "The update check failed (continuing): $($_.Exception.Message)" -Component 'Update'
        $state['LastCheckUtc'] = [DateTime]::UtcNow.ToString('o')
        $state['LastCheckResult'] = 'Error'
        $null = Write-UpdateState -Path $statePath -State $state
        return $null
    }
}

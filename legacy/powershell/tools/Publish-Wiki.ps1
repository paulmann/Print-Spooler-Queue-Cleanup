<#
.SYNOPSIS
    Publishes docs/wiki/ to the repository's GitHub wiki.

.DESCRIPTION
    A GitHub wiki is a SEPARATE git repository (<repo>.wiki.git) with a flat file
    layout: every page is a top-level .md file whose name becomes the page title
    (hyphens become spaces). There is no folder structure and no push API, so
    publishing is a clone-modify-push cycle against that repository.

    This script performs that cycle safely:

      * clones the existing wiki (or initialises an empty one on first publish)
      * copies docs/wiki/*.md across, preserving _Sidebar.md
      * rewrites relative links so they resolve in the wiki's flat namespace
      * commits and pushes only when something actually changed
      * supports -WhatIf to preview the diff without touching the remote

    Credentials come from git itself (credential helper, SSH key, or an
    https remote with a PAT). The script never handles a token directly.

.PARAMETER WikiUrl
    The wiki repository URL. Defaults to the project wiki derived from the
    constants in the built artifact.

.PARAMETER WhatIf
    Show what would be copied and committed; do not push.

.PARAMETER Message
    Commit message. Defaults to a version-stamped summary.

.EXAMPLE
    ./tools/Publish-Wiki.ps1 -WhatIf
.EXAMPLE
    ./tools/Publish-Wiki.ps1
.EXAMPLE
    ./tools/Publish-Wiki.ps1 -WikiUrl git@github.com:paulmann/Print-Spooler-Queue-Cleanup.wiki.git

.NOTES
    First-time publishing: GitHub creates the wiki repository only after you
    create the first page through the web UI. Create any page (it can be empty),
    then run this script - it will overwrite Home.md.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [AllowNull()][AllowEmptyString()][string]$WikiUrl,
    [AllowNull()][AllowEmptyString()][string]$SourceDirectory,
    [AllowNull()][AllowEmptyString()][string]$Message,
    [AllowNull()][AllowEmptyString()][string]$WorkDirectory
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot

if ([string]::IsNullOrWhiteSpace($SourceDirectory)) { $SourceDirectory = Join-Path $root 'docs/wiki' }
if (-not (Test-Path -LiteralPath $SourceDirectory -PathType Container)) {
    throw "Wiki source directory not found: $SourceDirectory"
}

if ([string]::IsNullOrWhiteSpace($WikiUrl)) {
    # Derive from the origin remote so a fork publishes to the fork.
    $origin = $null
    try {
        Push-Location $root
        $origin = (& git remote get-url origin 2>$null)
    }
    finally { Pop-Location }
    if ([string]::IsNullOrWhiteSpace($origin)) {
        $WikiUrl = 'https://github.com/paulmann/Print-Spooler-Queue-Cleanup.wiki.git'
        Write-Warning "No git origin remote found; defaulting to $WikiUrl"
    }
    elseif ($origin -match '^(git@github\.com):(.+?)(?:\.git)?$') {
        $WikiUrl = "{0}:{1}.wiki.git" -f $Matches[1], $Matches[2]
    }
    elseif ($origin -match '^https://github\.com/(.+?)(?:\.git)?$') {
        $WikiUrl = "https://github.com/$($Matches[1]).wiki.git"
    }
    else {
        throw "Could not derive a wiki URL from origin '$origin'. Pass -WikiUrl explicitly."
    }
}

$pages = @(Get-ChildItem -LiteralPath $SourceDirectory -Filter '*.md' -File | Sort-Object Name)
if ($pages.Count -eq 0) { throw "No .md pages found in $SourceDirectory" }

# The version, for the commit message and for the sidebar footer.
$version = '0.0.0'
$constants = Join-Path $root 'src/Shared/00-Constants.ps1'
if (Test-Path -LiteralPath $constants) {
    $m = [regex]::Match((Get-Content -LiteralPath $constants -Raw), "\`$Script:ScriptVersion\s+=\s+'([^']+)'")
    if ($m.Success) { $version = $m.Groups[1].Value }
}

if ([string]::IsNullOrWhiteSpace($Message)) {
    $Message = "docs(wiki): sync from docs/wiki at v$version"
}

Write-Host ''
Write-Host ("Wiki source : {0}" -f $SourceDirectory) -ForegroundColor Cyan
Write-Host ("Wiki remote : {0}" -f $WikiUrl) -ForegroundColor Cyan
Write-Host ("Pages       : {0}" -f $pages.Count) -ForegroundColor Cyan
Write-Host ("Version     : {0}" -f $version) -ForegroundColor Cyan
Write-Host ''

# ---- validate internal links resolve to real pages --------------------------
$pageNames = @($pages | ForEach-Object { [IO.Path]::GetFileNameWithoutExtension($_.Name) })
$broken = New-Object System.Collections.Generic.List[string]
foreach ($p in $pages) {
    $text = Get-Content -LiteralPath $p.FullName -Raw
    foreach ($m in [regex]::Matches($text, '\]\((?!https?:|mailto:|#)([^)#]+?)(?:#[^)]*)?\)')) {
        $target = $m.Groups[1].Value
        if ($target -notmatch '\.md$') { continue }
        $name = [IO.Path]::GetFileNameWithoutExtension($target)
        # Links to repository files (README, LICENSE, CHANGELOG) are absolute in
        # the wiki and legitimately point outside the page set.
        if ($target -match '^\.\./' -or $target -match '^/') { continue }
        if ($pageNames -notcontains $name) { $broken.Add(("{0} -> {1}" -f $p.Name, $target)) }
    }
}
if ($broken.Count -gt 0) {
    Write-Host 'Broken wiki links:' -ForegroundColor Red
    $broken | Select-Object -Unique | ForEach-Object { Write-Host ("   {0}" -f $_) -ForegroundColor Yellow }
    throw "$($broken.Count) wiki link(s) point at pages that do not exist."
}
Write-Host ("Link check: all {0} internal links resolve." -f $pages.Count) -ForegroundColor Green

# ---- stage the pages --------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($WorkDirectory)) {
    $WorkDirectory = Join-Path ([IO.Path]::GetTempPath()) ("psq-wiki-{0}" -f [Guid]::NewGuid().ToString('N'))
}
$null = New-Item -ItemType Directory -Path $WorkDirectory -Force

$cloneOk = $true
try {
    if (Test-Path -LiteralPath (Join-Path $WorkDirectory '.git')) {
        Push-Location $WorkDirectory
        try { & git fetch origin 2>&1 | Out-Null; & git reset --hard origin/master 2>&1 | Out-Null }
        finally { Pop-Location }
    }
    else {
        & git clone $WikiUrl $WorkDirectory 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0) { $cloneOk = $false }
    }
}
catch { $cloneOk = $false }

if (-not $cloneOk) {
    Write-Warning "Could not clone '$WikiUrl' (first publish, or no credentials). Initialising a fresh wiki repository instead."
    if (Test-Path -LiteralPath $WorkDirectory) { Remove-Item -LiteralPath $WorkDirectory -Recurse -Force -ErrorAction SilentlyContinue }
    $null = New-Item -ItemType Directory -Path $WorkDirectory -Force
    Push-Location $WorkDirectory
    try {
        & git init -q -b master 2>&1 | Out-Null
        & git remote add origin $WikiUrl 2>&1 | Out-Null
    }
    finally { Pop-Location }
}

# ---- copy pages, normalising links for the wiki's flat namespace ------------
$copied = 0
foreach ($p in $pages) {
    $text = Get-Content -LiteralPath $p.FullName -Raw

    # In the wiki every page is a sibling, so a bare "Page-Name" link works.
    # Repository-file links must become absolute github.com URLs, because the
    # wiki repository does not contain the project files.
    $text = [regex]::Replace($text, '\]\((?!https?:|mailto:|#|\.\./|/)([^)#]+?\.(?:ps1|bat|psd1|psm1|yml|txt|zip))(#[^)]*)?\)', {
            param($m)
            '](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/blob/main/{0}{1})' -f $m.Groups[1].Value, $m.Groups[2].Value
        })

    # ../wiki/X.md and ../X.md references collapse to the flat page name.
    $text = [regex]::Replace($text, '\]\((?:\.\./)+(?:wiki/)?([A-Za-z0-9_-]+)\.md(#[^)]*)?\)', ']($1$2)')

    $dest = Join-Path $WorkDirectory $p.Name
    Set-Content -LiteralPath $dest -Value $text -Encoding UTF8 -NoNewline
    $copied++
}
Write-Host ("Staged {0} page(s)." -f $copied) -ForegroundColor Green

# A wiki without a Home page renders an empty front page.
if (-not (Test-Path -LiteralPath (Join-Path $WorkDirectory 'Home.md'))) {
    throw 'No Home.md in the wiki source - GitHub would show an empty front page.'
}

# ---- commit and push --------------------------------------------------------
Push-Location $WorkDirectory
try {
    # Under -WhatIf the page files were never written, so `git status` would
    # report "no changes" and the run would misleadingly claim the wiki is
    # already up to date. Report the intent instead and stop before any git
    # write operation.
    if ($WhatIfPreference) {
        Write-Host ''
        Write-Host ("WhatIf: would publish {0} page(s) to {1}" -f $copied, $WikiUrl) -ForegroundColor Yellow
        Write-Host '        Nothing was committed or pushed.' -ForegroundColor Yellow
        return
    }

    & git add -A 2>&1 | Out-Null
    $status = (& git status --porcelain 2>&1)
    if ([string]::IsNullOrWhiteSpace(($status -join ''))) {
        Write-Host 'No changes to publish - the wiki is already up to date.' -ForegroundColor Green
        return
    }

    Write-Host ''
    Write-Host 'Changes to publish:' -ForegroundColor Cyan
    $status | ForEach-Object { Write-Host ("   {0}" -f $_) }

    if (-not $PSCmdlet.ShouldProcess($WikiUrl, "Push $($copied) wiki page(s)")) { return }

    & git -c user.name='Print Spooler Queue Cleanup' -c user.email='Mikhail@Deynekin.com' commit -q -m $Message 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'git commit failed.' }

    & git push -q origin master 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "git push failed. Check that the wiki is enabled on the repository and that your credentials can write to '$WikiUrl'."
    }
    Write-Host ''
    Write-Host "Published $($copied) page(s) to $WikiUrl" -ForegroundColor Green
}
finally {
    Pop-Location
    if ($WorkDirectory -like (Join-Path ([IO.Path]::GetTempPath()) '*')) {
        Remove-Item -LiteralPath $WorkDirectory -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# ============================================================================
#  General utilities: versions, paths, time, hashing, validation
# ============================================================================

function ConvertTo-SemVer {
    <#
    .SYNOPSIS
        Parses "v4.0.0", "4.0", "4.0.0-beta.1+build" into a comparable object.
    .OUTPUTS
        PSCustomObject { Version, Label, IsPrerelease, Raw, Normalized }; $null when unparseable.
    .NOTES
        [Version] only understands numeric dot-separated parts, so pre-release and
        build metadata are split off and compared separately.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory, Position = 0)][AllowNull()][AllowEmptyString()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $raw = $Value.Trim().TrimStart('v', 'V')

    $core = $raw
    $label = ''
    $plusIdx = $core.IndexOf('+')
    if ($plusIdx -ge 0) { $core = $core.Substring(0, $plusIdx) }
    $dashIdx = $core.IndexOf('-')
    if ($dashIdx -ge 0) {
        $label = $core.Substring($dashIdx + 1)
        $core = $core.Substring(0, $dashIdx)
    }

    $parts = @($core.Split('.') | Where-Object { $_ -ne '' })
    if ($parts.Count -eq 0) { return $null }
    foreach ($p in $parts) { if ($p -notmatch '^\d+$') { return $null } }
    while ($parts.Count -lt 4) { $parts += '0' }
    if ($parts.Count -gt 4) { $parts = $parts[0..3] }

    try { $ver = [Version]::new([int]$parts[0], [int]$parts[1], [int]$parts[2], [int]$parts[3]) }
    catch { return $null }

    return [pscustomobject]@{
        Version      = $ver
        Label        = $label
        IsPrerelease = (-not [string]::IsNullOrEmpty($label))
        Raw          = $Value.Trim()
        Normalized   = $ver.ToString()
    }
}

function Compare-ProductVersion {
    <#
    .SYNOPSIS
        -1 when $Reference < $Difference, 0 when equal, +1 when greater.
    .NOTES
        Test-enforced against a table of known-tricky pairs.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param(
        # [AllowEmptyString()] matters: without it a Mandatory [string] parameter
        # rejects '' at bind time, so an absent version tag from GitHub would throw
        # instead of being handled as "unknown".
        [Parameter(Mandatory, Position = 0)][AllowNull()][AllowEmptyString()][string]$Reference,
        [Parameter(Mandatory, Position = 1)][AllowNull()][AllowEmptyString()][string]$Difference
    )
    $a = ConvertTo-SemVer -Value $Reference
    $b = ConvertTo-SemVer -Value $Difference
    if (-not $a -and -not $b) { return 0 }
    if (-not $a) { return -1 }
    if (-not $b) { return 1 }

    $c = $a.Version.CompareTo($b.Version)
    if ($c -ne 0) { return [int][Math]::Sign($c) }

    if ($a.IsPrerelease -and -not $b.IsPrerelease) { return -1 }
    if (-not $a.IsPrerelease -and $b.IsPrerelease) { return 1 }
    return [int][Math]::Sign([string]::CompareOrdinal($a.Label, $b.Label))
}

function Test-NewerVersion {
    <#
    .SYNOPSIS
        True when $Candidate is strictly newer than $Current.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory, Position = 0)][AllowNull()][AllowEmptyString()][string]$Current,
        [Parameter(Mandatory, Position = 1)][AllowNull()][AllowEmptyString()][string]$Candidate,
        [switch]$IncludePrerelease
    )
    $c = ConvertTo-SemVer -Value $Candidate
    if (-not $c) { return $false }
    if ($c.IsPrerelease -and -not $IncludePrerelease) { return $false }
    return ((Compare-ProductVersion -Reference $Current -Difference $Candidate) -lt 0)
}

function ConvertFrom-AnyTimeSpan {
    <#
    .SYNOPSIS
        Accepts [TimeSpan], a number (seconds) or a string ("90s","2h","3d","01:30:00")
        and returns a [TimeSpan]. Throws on garbage.
    .DESCRIPTION
        Powers -OlderThan / -NewerThan, which operators expect to type as
        "-OlderThan 2h" rather than "[TimeSpan]::FromHours(2)".
    #>
    [CmdletBinding()]
    [OutputType([TimeSpan])]
    param([Parameter(Mandatory, Position = 0)][AllowNull()]$Value)

    if ($null -eq $Value) { return [TimeSpan]::Zero }
    if ($Value -is [TimeSpan]) { return $Value }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) {
        return [TimeSpan]::FromSeconds([double]$Value)
    }
    $s = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($s)) { return [TimeSpan]::Zero }

    if ($s -match '^(?<num>\d+(?:\.\d+)?)\s*(?<unit>ms|millisecond(s)?|s|sec(ond)?(s)?|m|min(ute)?(s)?|h|hr|hour(s)?|d|day(s)?|w|wk|week(s)?)$') {
        $num = [double]$Matches['num']
        $unit = $Matches['unit'].ToLowerInvariant()
        # FOUR PowerShell traps avoided here, each of which produced a real defect:
        #   1. `return switch ...` is a SYNTAX ERROR - switch is a statement, not an
        #      expression, so it cannot be the operand of return.
        #   2. switch does NOT stop at the first match, and -regex is a substring
        #      match, so unanchored patterns made 'ms' satisfy BOTH the millisecond
        #      and the minute clause and return an ARRAY of two TimeSpans.
        #   3. A '$' immediately before a closing single quote forms the special
        #      variable $' and silently swallows the terminator.
        #   4. The result is wrapped in @() and indexed so a multi-match can never
        #      leak an array into a [TimeSpan] return value again.
        $spans = @(switch -regex ($unit) {
                '^(?:ms|millisecond|milliseconds)$' { [TimeSpan]::FromMilliseconds($num) }
                '^(?:w|wk|week|weeks)$' { [TimeSpan]::FromDays($num * 7) }
                '^(?:d|day|days)$' { [TimeSpan]::FromDays($num) }
                '^(?:h|hr|hrs|hour|hours)$' { [TimeSpan]::FromHours($num) }
                '^(?:m|min|mins|minute|minutes)$' { [TimeSpan]::FromMinutes($num) }
                '^(?:s|sec|secs|second|seconds)$' { [TimeSpan]::FromSeconds($num) }
            })
        if ($spans.Count -eq 0) { $spans = @([TimeSpan]::FromSeconds($num)) }
        return [TimeSpan]$spans[0]
    }

    # A bare number means SECONDS, which is what the help promises. It must be
    # handled before [TimeSpan]::TryParse, because .NET interprets a bare number
    # as DAYS - silently turning "-OlderThan 30" into a 30-day filter.
    if ($s -match '^[-+]?\d+(?:\.\d+)?$') { return [TimeSpan]::FromSeconds([double]$s) }

    # Fall back to the standard TimeSpan parser ("1.02:03:04").
    $ts = [TimeSpan]::Zero
    if ([TimeSpan]::TryParse($s, [ref]$ts)) { return $ts }
    throw ("Cannot interpret '{0}' as a time span. Use e.g. 30s, 15m, 2h, 7d or '01:30:00'." -f $s)
}

function Get-SafeFileNameStamp {
    <#
    .SYNOPSIS
        Culture-invariant timestamp suitable for file names.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [AllowNull()][DateTime]$When,
        [string]$Format = 'yyyyMMdd-HHmmss'
    )
    $d = if ($When -eq [DateTime]::MinValue) { Get-Date } else { $When }
    return $d.ToString($Format, [Globalization.CultureInfo]::InvariantCulture)
}

function ConvertTo-SafeFileName {
    <#
    .SYNOPSIS
        Strips characters that are illegal in Windows file names.
    .NOTES
        Uses the Windows illegal set explicitly rather than
        [IO.Path]::GetInvalidFileNameChars(), because that API reflects the HOST
        platform: on Linux it omits '\', ':' and '"', so a name sanitised there
        would still be invalid on the Windows machines this tool targets.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, Position = 0)][AllowNull()][AllowEmptyString()][string]$Value,
        [string]$Replacement = '_',
        [int]$MaxLength = 120
    )
    if ([string]::IsNullOrWhiteSpace($Value)) { return 'unnamed' }

    $invalid = @([IO.Path]::GetInvalidFileNameChars()) + @('\', '/', ':', '*', '?', '"', '<', '>', '|')
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Value.ToCharArray()) {
        $cs = [string]$ch
        if ($invalid -contains $cs -or $cs -eq ' ') { [void]$sb.Append($Replacement) }
        else { [void]$sb.Append($cs) }
    }
    $out = $sb.ToString()
    $out = $out -replace "$([regex]::Escape($Replacement)){2,}", $Replacement
    $out = $out.Trim($Replacement.ToCharArray())
    if ([string]::IsNullOrWhiteSpace($out)) { $out = 'unnamed' }
    if ($out.Length -gt $MaxLength) { $out = $out.Substring(0, $MaxLength) }
    return $out
}

function Resolve-DirectoryPath {
    <#
    .SYNOPSIS
        Expands environment variables, resolves to a full path, optionally creates.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, Position = 0)][AllowNull()][AllowEmptyString()][string]$Path,
        [switch]$Create
    )
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $expanded = [Environment]::ExpandEnvironmentVariables($Path.Trim()).Trim('"')
    try { $full = [IO.Path]::GetFullPath($expanded) }
    catch { $full = $expanded }
    if ($Create -and -not (Test-Path -LiteralPath $full)) {
        New-Item -ItemType Directory -Path $full -Force -ErrorAction Stop | Out-Null
    }
    return $full
}

function Test-DirectoryWritable {
    <#
    .SYNOPSIS
        Real write probe (create and delete a temp file), not an ACL guess.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory, Position = 0)][AllowNull()][AllowEmptyString()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return $false }
    $probe = Join-Path $Path (".psq-write-probe-{0}" -f ([Guid]::NewGuid().ToString('N')))
    try {
        Set-Content -LiteralPath $probe -Value 'probe' -Encoding ASCII -ErrorAction Stop
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
        return $true
    }
    catch {
        return $false
    }
}

function Get-FileHashSafe {
    <#
    .SYNOPSIS
        Hash of a file, or $null instead of throwing.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, Position = 0)][AllowNull()][AllowEmptyString()][string]$Path,
        [ValidateSet('SHA256', 'SHA512', 'MD5')][string]$Algorithm = 'SHA256'
    )
    try {
        if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
        return (Get-FileHash -LiteralPath $Path -Algorithm $Algorithm -ErrorAction Stop).Hash
    }
    catch {
        return $null
    }
}

function Test-HashMatch {
    <#
    .SYNOPSIS
        Hash comparison that is case-insensitive and tolerant of separators.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        # [AllowEmptyString()] is required: a Mandatory [string] rejects '' at bind
        # time, before the guard below could return $false. An empty expected hash
        # is a legitimate "no checksum was published" case.
        [Parameter(Mandatory, Position = 0)][AllowNull()][AllowEmptyString()][string]$Expected,
        [Parameter(Mandatory, Position = 1)][AllowNull()][AllowEmptyString()][string]$Actual
    )
    if ([string]::IsNullOrWhiteSpace($Expected) -or [string]::IsNullOrWhiteSpace($Actual)) { return $false }
    $e = ($Expected -replace '[^0-9A-Fa-f]', '').ToLowerInvariant()
    $a = ($Actual -replace '[^0-9A-Fa-f]', '').ToLowerInvariant()
    if ($e.Length -ne $a.Length -or $e.Length -eq 0) { return $false }
    return [string]::Equals($e, $a, [StringComparison]::Ordinal)
}

function ConvertTo-HashtableDeep {
    <#
    .SYNOPSIS
        Recursively converts PSCustomObject trees into plain hashtables.
    .DESCRIPTION
        Required because objects sent to a remote session must survive CLIXML
        serialization; hashtables of scalars always do, while PSObject property
        bags can arrive with type decorations that make member access unreliable
        on the far side.
    #>
    [CmdletBinding()]
    param([AllowNull()]$InputObject, [int]$Depth = 0)

    if ($Depth -gt 12) { return $null }
    if ($null -eq $InputObject) { return $null }

    if ($InputObject -is [System.Collections.IDictionary]) {
        $out = @{}
        foreach ($k in @($InputObject.Keys)) { $out[[string]$k] = ConvertTo-HashtableDeep -InputObject $InputObject[$k] -Depth ($Depth + 1) }
        return $out
    }
    if (($InputObject -is [System.Collections.IEnumerable]) -and ($InputObject -isnot [string])) {
        $list = New-Object System.Collections.Generic.List[object]
        foreach ($item in $InputObject) { $list.Add((ConvertTo-HashtableDeep -InputObject $item -Depth ($Depth + 1))) }
        return $list.ToArray()
    }
    if ($InputObject -is [pscustomobject]) {
        $out = @{}
        foreach ($p in $InputObject.PSObject.Properties) { $out[$p.Name] = ConvertTo-HashtableDeep -InputObject $p.Value -Depth ($Depth + 1) }
        return $out
    }
    if ($InputObject -is [TimeSpan]) { return $InputObject.TotalSeconds }
    if ($InputObject -is [DateTime]) { return $InputObject.ToString('o', [Globalization.CultureInfo]::InvariantCulture) }
    if ($InputObject -is [Enum]) { return [string]$InputObject }
    return $InputObject
}

function New-ResultObject {
    <#
    .SYNOPSIS
        Factory for the per-computer result record. This shape is a public contract.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$ComputerName,
        [string]$Status = 'Unknown'
    )
    return [pscustomobject]([ordered]@{
            ComputerName         = $ComputerName
            Status               = $Status
            StartedUtc           = [DateTime]::UtcNow
            FinishedUtc          = $null
            DurationMs           = [long]0
            IsLocal              = $false
            SpoolDirectory       = $null
            SpoolDirectorySource = $null
            ServiceWasRunning    = $null
            ServiceIsRunning     = $null
            ServiceStopMs        = [long]0
            ServiceStartMs       = [long]0
            DependentServices    = @()
            JobsFound            = 0
            JobsSelected         = 0
            JobsRemoved          = 0
            FilesFound           = 0
            FilesSelected        = 0
            FilesDeleted         = 0
            FilesFailed          = 0
            FilesLocked          = 0
            BytesFound           = [long]0
            BytesDeleted         = [long]0
            FilesAfter           = 0
            ResurrectedFiles     = 0
            BackupPath           = $null
            BackupHash           = $null
            DryRun               = $false
            RepairPerformed      = $false
            DiagnosticsCollected = $false
            Warnings             = (New-Object System.Collections.Generic.List[string])
            Errors               = (New-Object System.Collections.Generic.List[string])
            Jobs                 = @()
            RemovedFiles         = @()
            FailedFiles          = @()
            OrphanedFiles        = @()
            ServiceTrace         = @()
            Diagnostics          = $null
            Mode                 = $null
            ExitCode             = $null
        })
}

function Add-ResultWarning {
    <#
    .SYNOPSIS
        Records a warning on a result object and in the log.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Result, [Parameter(Mandatory)][AllowEmptyString()][string]$Message)
    if ($Result.Warnings -is [System.Collections.Generic.List[string]]) { $Result.Warnings.Add($Message) | Out-Null }
    Write-LogWarn $Message -Component 'Cleanup' -Target $Result.ComputerName
}

function Add-ResultError {
    <#
    .SYNOPSIS
        Records an error on a result object and in the log.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Result, [Parameter(Mandatory)][AllowEmptyString()][string]$Message)
    if ($Result.Errors -is [System.Collections.Generic.List[string]]) { $Result.Errors.Add($Message) | Out-Null }
    Write-LogError $Message -Component 'Cleanup' -Target $Result.ComputerName
}

function Test-IsLocalTarget {
    <#
    .SYNOPSIS
        True for localhost / . / 127.0.0.1 / ::1 / the local NetBIOS and DNS names.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory, Position = 0)][AllowNull()][AllowEmptyString()][string]$ComputerName)

    if ([string]::IsNullOrWhiteSpace($ComputerName)) { return $true }
    $name = $ComputerName.Trim()
    # TrimEnd('.') normalises FQDNs (SRV01.corp.local. -> SRV01.corp.local), but it
    # would also reduce the literal local alias '.' to an empty string and make it
    # unmatchable, so that case is guarded explicitly.
    if ($name -ne '.' -and $name -ne '..') { $name = $name.TrimEnd('.') }
    foreach ($local in $Script:RuntimeState.LocalNames) {
        if ([string]::Equals($name, $local, [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    if ($name -match '^([^.]+)\.') {
        $netbios = $Matches[1]
        foreach ($local in $Script:RuntimeState.LocalNames) {
            if ([string]::Equals($netbios, $local, [StringComparison]::OrdinalIgnoreCase)) { return $true }
        }
    }
    return $false
}

function Get-LocalComputerNames {
    <#
    .SYNOPSIS
        Every name that refers to "this machine", used by Test-IsLocalTarget.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param()
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($n in @('localhost', '.', '127.0.0.1', '::1', '(local)', $env:COMPUTERNAME, $env:HOSTNAME)) {
        if (-not [string]::IsNullOrWhiteSpace($n)) { $names.Add($n) }
    }
    try { $names.Add([System.Net.Dns]::GetHostName()) }
    catch { /* the DNS name is optional; the other aliases are sufficient */ }
    return @($names | Select-Object -Unique)
}

function BlankValueMessage {
    <#
    .SYNOPSIS
        Builds the rejection message for a blank selection/path value.
    .DESCRIPTION
        The wording matters: the operator has to understand that the run was
        REFUSED, not silently adjusted, and why a blank filter is dangerous
        rather than merely useless.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Kind
    )
    return ("-{0} was given {1}. A blank selection filter is falsy in PowerShell, so it would be IGNORED rather than matched - silently widening the run to EVERY printer and every job. Specify a real value, or omit -{0} entirely." -f $Name, $Kind)
}

function Test-ParameterCombination {
    <#
    .SYNOPSIS
        Cross-parameter validation that ParameterSetName cannot express.
    .DESCRIPTION
        Returns human-readable conflict messages (empty = valid). Keeping the rules
        here makes them unit-testable without invoking the whole cleanup pipeline.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param([Parameter(Mandatory)][AllowNull()][System.Collections.IDictionary]$Bound)

    $problems = New-Object System.Collections.Generic.List[string]
    if (-not $Bound) { return $problems.ToArray() }

    if ($Bound.Contains('Restore') -and $Bound.Contains('Repair')) {
        $problems.Add('-Restore and -Repair are mutually exclusive.')
    }
    if ($Bound.Contains('Restore') -and ($Bound.Contains('PrinterName') -or $Bound.Contains('JobId') -or $Bound.Contains('OlderThan'))) {
        $problems.Add('-Restore cannot be combined with selection filters (-PrinterName/-JobId/-OlderThan).')
    }
    if ($Bound.Contains('Diagnostic') -and $Bound.Contains('Repair')) {
        $problems.Add('-Diagnostic is read-only and cannot be combined with -Repair.')
    }
    if ($Bound.Contains('Diagnostic') -and ($Bound.Contains('PrinterName') -or $Bound.Contains('JobId'))) {
        $problems.Add('-Diagnostic reports everything; remove the selection filters or drop -Diagnostic.')
    }
    if ($Bound.Contains('BackupPath') -and -not $Bound.Contains('Backup')) {
        $problems.Add('-BackupPath requires -Backup.')
    }
    if ($Bound.Contains('OlderThan') -and $Bound.Contains('NewerThan')) {
        $older = ConvertFrom-AnyTimeSpan -Value $Bound['OlderThan']
        $newer = ConvertFrom-AnyTimeSpan -Value $Bound['NewerThan']
        if ($older -le $newer) { $problems.Add('-OlderThan must be greater than -NewerThan, otherwise nothing can ever match.') }
    }
    if ($Bound.Contains('ThrottleLimit') -and [int]$Bound['ThrottleLimit'] -gt 1 -and -not (Test-PowerShellSevenFeatures)) {
        $problems.Add("-ThrottleLimit > 1 needs PowerShell 7 for ForEach-Object -Parallel. Current host: $($PSVersionTable.PSVersion). (A runspace pool is used automatically, so this is advisory.)")
    }
    if ($Bound.Contains('EventLog') -and -not (Test-RunningOnWindows)) {
        $problems.Add('-EventLog is only available on Windows.')
    }
    if ($Bound.Contains('AutoUpdate') -and $Bound.Contains('CheckUpdate')) {
        $problems.Add('-AutoUpdate already implies an update check; remove -CheckUpdate.')
    }
    if ($Bound.Contains('Update') -and $Bound.Contains('ComputerName')) {
        $problems.Add('-Update updates this tool only; it cannot be combined with -ComputerName.')
    }
    if ($Bound.Contains('SelfElevate') -and $Bound.Contains('NoElevate')) {
        $problems.Add('-SelfElevate and -NoElevate are mutually exclusive.')
    }
    if ($Bound.Contains('SaveConfig') -and -not $Bound.Contains('ConfigPath')) {
        $problems.Add('-SaveConfig requires -ConfigPath <file> to write to.')
    }
    if ($Bound.Contains('MachineReadableJson') -and $Bound.Contains('ExportFormat')) {
        $problems.Add('-MachineReadableJson writes JSON to stdout; use -ExportPath instead of -ExportFormat.')
    }
    if ($Bound.Contains('WatchdogMinutes') -and [int]$Bound['WatchdogMinutes'] -lt 0) {
        $problems.Add('-WatchdogMinutes cannot be negative (0 disables the watchdog).')
    }
    if ($Bound.Contains('JobId') -and $Bound.Contains('PrinterName')) {
        $problems.Add('-JobId already identifies a job exactly; combining it with -PrinterName is ambiguous. Use one of them.')
    }
    if ($Bound.Contains('Repair') -and $Bound.Contains('PurgeDriverStore') -and -not $Bound.Contains('Force')) {
        $problems.Add('-PurgeDriverStore removes driver packages and therefore requires -Force as an explicit acknowledgement.')
    }

    # ---- a blank value on a selection or path parameter is a SAFETY error -----
    # An empty string is FALSY in PowerShell. The payload gates every filter on
    #     if ($Options['PrinterName'] -and @($Options['PrinterName']).Count -gt 0)
    # so a filter of '' does not narrow the selection at all - the guard fails and
    # the run silently becomes NON-selective, cleaning EVERY printer instead of the
    # one the operator named. The mirror case is worse: a blank
    # -ExcludePrinterName silently removes a protection the operator asked for.
    #
    # This is reachable from the .bat launcher, where `/printer` typed without a
    # value forwards -PrinterName "". Reject it loudly instead of letting a typo
    # widen the blast radius of a destructive run. Exit code 3 (BadArguments).
    foreach ($name in $Script:BlankValueIsUnsafeParameters) {
        if (-not $Bound.Contains($name)) { continue }
        $val = $Bound[$name]
        if ($null -eq $val) {
            $problems.Add((BlankValueMessage -Name $name -Kind 'no value'))
            continue
        }
        $items = @($val)
        if ($items.Count -eq 0) {
            $problems.Add((BlankValueMessage -Name $name -Kind 'an empty list'))
            continue
        }
        $allBlank = $true
        foreach ($it in $items) {
            if (-not [string]::IsNullOrWhiteSpace([string]$it)) { $allBlank = $false; break }
        }
        if ($allBlank) { $problems.Add((BlankValueMessage -Name $name -Kind "a blank value ('')")) }
    }

    return $problems.ToArray()
}

function Get-EnvironmentConfiguration {
    <#
    .SYNOPSIS
        Reads PSQ_* environment variables into a hashtable with the prefix stripped.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param([string]$Prefix = $Script:ConfigEnvPrefix)

    $map = @{}
    try {
        foreach ($entry in [Environment]::GetEnvironmentVariables().GetEnumerator()) {
            $key = [string]$entry.Key
            if ($key.StartsWith($Prefix, [StringComparison]::OrdinalIgnoreCase)) {
                $map[$key.Substring($Prefix.Length)] = [string]$entry.Value
            }
        }
    }
    catch {
        # An unreadable environment block must not stop the run.
    }
    return $map
}

function ConvertFrom-EnvironmentBoolean {
    <#
    .SYNOPSIS
        Interprets the usual spellings of a boolean environment variable.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Position = 0)][AllowNull()][AllowEmptyString()][string]$Value,
        [Parameter(Position = 1)][bool]$Default = $false
    )
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Default }
    return ($Value.Trim().ToLowerInvariant() -in @('1', 'true', 'yes', 'y', 'on', 'enabled'))
}

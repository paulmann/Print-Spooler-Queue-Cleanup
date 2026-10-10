# ============================================================================
#  Logging subsystem
# ----------------------------------------------------------------------------
#  FINDING A14 (fixed): v3 wrote a single ever-growing file next to the script,
#  which fails outright when the script lives somewhere read-only
#  (C:\Program Files, a locked-down SMB share, a mounted ISO). v4:
#
#    * probes several candidate roots and uses the FIRST WRITABLE one
#    * supports an explicit -LogPath, creating the parent directory as needed
#    * rotates by size (default 10 MB) keeping N generations (.1 .. .9)
#    * applies a retention policy (default 30 days) on every run
#    * writes a pipe-delimited, locale-independent, UTC-stamped record format
#      that parses identically in PowerShell, grep, Splunk or ELK
#    * can emit JSON Lines instead (-JsonLog) for log shippers
#    * keeps an in-memory ring buffer for the diagnostics bundle and for
#      post-mortem analysis when the file itself is unusable
#    * is thread-safe, because PowerShell 7 -Parallel workers log concurrently
#
#  Record format (default):
#    2026-10-10T12:34:56.789Z|INFO    |Spooler|SRV01|Message text
#  Record format (JSON Lines):
#    {"ts":"...","level":"INFO","component":"Spooler","target":"SRV01","msg":"..."}
# ============================================================================

function Resolve-LogTarget {
    <#
    .SYNOPSIS
        Chooses the log file path: explicit -LogPath, then the first writable root.
    .OUTPUTS
        PSCustomObject { Path, Source, Writable, Error }
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyString()][string]$ExplicitPath,
        [string[]]$Roots = $Script:LogCandidateRoots,
        [string]$FileNamePrefix = $Script:LogFileNamePrefix,
        [switch]$SessionStamped
    )

    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $fileName = if ($SessionStamped) { "{0}_{1}.log" -f $FileNamePrefix, $stamp } else { "{0}.log" -f $FileNamePrefix }

    if (-not [string]::IsNullOrWhiteSpace($ExplicitPath)) {
        $full = $null
        try {
            $full = [IO.Path]::GetFullPath($ExplicitPath)
            $dir = [IO.Path]::GetDirectoryName($full)
            if ($dir -and -not (Test-Path -LiteralPath $dir)) {
                New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null
            }
            # Prove writability by opening the file, not by guessing from ACLs.
            # The probe writes no bytes, so it cannot leave a blank record that
            # would break a strict line-oriented log parser.
            $probe = [IO.File]::Open($full, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
            $probe.Dispose()
            return [pscustomobject]@{ Path = $full; Source = 'Parameter'; Writable = $true; Error = $null }
        }
        catch {
            return [pscustomobject]@{ Path = $full; Source = 'Parameter'; Writable = $false; Error = $_.Exception.Message }
        }
    }

    foreach ($root in $Roots) {
        $dir = switch ($root) {
            'ProgramData' { if ($env:ProgramData) { Join-Path $env:ProgramData $Script:ProductFamily } }
            'LocalAppData' { if ($env:LOCALAPPDATA) { Join-Path $env:LOCALAPPDATA $Script:ProductFamily } }
            'ScriptDirectory' { if ($Script:RuntimeState.ScriptDirectory) { Join-Path $Script:RuntimeState.ScriptDirectory 'Logs' } }
            'Temp' { Join-Path ([IO.Path]::GetTempPath()) $Script:ProductFamily }
            default { $null }
        }
        if ([string]::IsNullOrWhiteSpace($dir)) { continue }
        try {
            if (-not (Test-Path -LiteralPath $dir)) {
                New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null
            }
            $candidate = Join-Path $dir $fileName
            $probe = [IO.File]::Open($candidate, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
            $probe.Dispose()
            return [pscustomobject]@{ Path = $candidate; Source = $root; Writable = $true; Error = $null }
        }
        catch {
            # Read-only or denied: fall through to the next candidate root.
            continue
        }
    }

    return [pscustomobject]@{ Path = $null; Source = 'None'; Writable = $false; Error = 'No writable log location found.' }
}

function Write-LogEntryCore {
    <#
    .SYNOPSIS
        The one and only log sink. All console and file logging funnels through here.
    .NOTES
        Deliberately never throws: a logging failure must not abort a cleanup that
        is halfway through having the Spooler stopped (finding A21).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)][AllowEmptyString()][string]$Message,
        [ValidateSet('TRACE', 'DEBUG', 'INFO', 'SUCCESS', 'WARN', 'ERROR', 'CRITICAL')]
        [string]$Level = 'INFO',
        [string]$Component = 'Core',
        [AllowNull()][string]$Target,
        [AllowNull()]$Data,
        [switch]$NoConsole,
        [switch]$NoFile
    )

    $now = [DateTime]::UtcNow
    $state = Get-RuntimeState
    $minIdx = if ($state.LogLevelMap.Contains($Level)) { [int]$state.LogLevelMap[$Level] } else { 2 }
    $minThreshold = [int]$state.MinLogLevelIndex

    # ---- in-memory ring buffer (always captured; used by the diagnostics bundle)
    try {
        if ($state.CaptureBuffer) {
            $entryObj = [pscustomobject]@{ TimeUtc = $now; Level = $Level; Component = $Component; Target = $Target; Message = $Message }
            if ($state.BufferSync) {
                [System.Threading.Monitor]::Enter($state.BufferSync)
                try {
                    $state.Buffer.Add($entryObj) | Out-Null
                    while ($state.Buffer.Count -gt $state.BufferMax) { $state.Buffer.RemoveAt(0) }
                }
                finally { [System.Threading.Monitor]::Exit($state.BufferSync) }
            }
            else {
                $state.Buffer.Add($entryObj) | Out-Null
                while ($state.Buffer.Count -gt $state.BufferMax) { $state.Buffer.RemoveAt(0) }
            }
        }
    }
    catch {
        # The buffer is a convenience; losing an entry is not worth an exception.
    }

    # ---- file
    if (-not $NoFile -and $state.LogPath -and ($minIdx -ge $minThreshold)) {
        try {
            $line = if ($state.JsonLog) {
                $rec = [ordered]@{
                    ts        = $now.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
                    level     = $Level
                    component = $Component
                    target    = $Target
                    msg       = $Message
                }
                if ($null -ne $Data) { $rec['data'] = $Data }
                ($rec | ConvertTo-Json -Compress -Depth 6)
            }
            else {
                $parts = @(
                    $now.ToString('yyyy-MM-ddTHH:mm:ss.fffZ'),
                    $Level.PadRight(8).Substring(0, 8),
                    $Component,
                    $(if ([string]::IsNullOrWhiteSpace($Target)) { '-' } else { $Target }),
                    ($Message -replace '[\r\n]+', ' ')
                )
                ($parts -join '|')
            }
            if ($state.LogSync) {
                [System.Threading.Monitor]::Enter($state.LogSync)
                try { Add-Content -LiteralPath $state.LogPath -Value $line -Encoding UTF8 -ErrorAction Stop }
                finally { [System.Threading.Monitor]::Exit($state.LogSync) }
            }
            else {
                Add-Content -LiteralPath $state.LogPath -Value $line -Encoding UTF8 -ErrorAction Stop
            }
        }
        catch {
            if (-not $state.LogFailureWarned) {
                $state.LogFailureWarned = $true
                try { Write-Warning ("Logging disabled: cannot write to '{0}' ({1})." -f $state.LogPath, $_.Exception.Message) }
                catch { /* the logger must never throw: losing one message beats losing the run */ }
            }
        }
    }

    # ---- console
    if (-not $NoConsole -and ($minIdx -ge $minThreshold) -and -not $state.Quiet) {
        if ($state.MachineReadable) { return }   # stdout is reserved for JSON only
        try { Write-LogToConsole -Message $Message -Level $Level -Component $Component -Target $Target }
        catch { /* host rendering is best-effort; the file record is the source of truth */ }
    }
}

function Write-LogToConsole {
    <#
    .SYNOPSIS
        Renders one coloured, glyph-prefixed, timestamped line to the host.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Message,
        [string]$Level = 'INFO',
        [string]$Component = 'Core',
        [AllowNull()][string]$Target
    )

    $state = $Script:RuntimeState
    $glyphs = $state.Glyphs
    $colors = $state.Colors

    $glyphKey = switch ($Level) {
        'TRACE' { 'Trace' }
        'DEBUG' { 'Debug' }
        'INFO' { 'Info' }
        'SUCCESS' { 'Success' }
        'WARN' { 'Warning' }
        'ERROR' { 'Error' }
        'CRITICAL' { 'Critical' }
        default { 'Info' }
    }
    $colorKey = $glyphKey

    $glyph = if ($glyphs.Contains($glyphKey)) { $glyphs[$glyphKey] } else { '*' }
    $color = if (-not $state.NoColor -and $colors.Contains($colorKey)) { $colors[$colorKey] } else { $null }
    $stamp = (Get-Date).ToString('HH:mm:ss')
    $tgtPart = if ([string]::IsNullOrWhiteSpace($Target)) { '' } else { "[$Target] " }
    $prefix = if ($state.ShowTimestamp) { "$stamp " } else { '' }

    $line = "{0}{1} {2}{3}" -f $prefix, $glyph, $tgtPart, $Message

    if ($null -ne $color) { Write-Host $line -ForegroundColor $color }
    else { Write-Host $line }

    if ($state.VerboseTranscript) { Write-Verbose $line }
}

# ---------------------------------------------------------------------------
#  Level shorthands
# ---------------------------------------------------------------------------
function Write-LogInfo {
    [CmdletBinding()]
    param([Parameter(Mandatory, Position = 0)][AllowEmptyString()][string]$Message, [string]$Component = 'Core', [AllowNull()][string]$Target, $Data)
    Write-LogEntryCore -Message $Message -Level 'INFO' -Component $Component -Target $Target -Data $Data
}
function Write-LogSuccess {
    [CmdletBinding()]
    param([Parameter(Mandatory, Position = 0)][AllowEmptyString()][string]$Message, [string]$Component = 'Core', [AllowNull()][string]$Target, $Data)
    Write-LogEntryCore -Message $Message -Level 'SUCCESS' -Component $Component -Target $Target -Data $Data
}
function Write-LogWarn {
    [CmdletBinding()]
    param([Parameter(Mandatory, Position = 0)][AllowEmptyString()][string]$Message, [string]$Component = 'Core', [AllowNull()][string]$Target, $Data)
    Write-LogEntryCore -Message $Message -Level 'WARN' -Component $Component -Target $Target -Data $Data
}
function Write-LogError {
    [CmdletBinding()]
    param([Parameter(Mandatory, Position = 0)][AllowEmptyString()][string]$Message, [string]$Component = 'Core', [AllowNull()][string]$Target, $Data)
    Write-LogEntryCore -Message $Message -Level 'ERROR' -Component $Component -Target $Target -Data $Data
}
function Write-LogCritical {
    [CmdletBinding()]
    param([Parameter(Mandatory, Position = 0)][AllowEmptyString()][string]$Message, [string]$Component = 'Core', [AllowNull()][string]$Target, $Data)
    Write-LogEntryCore -Message $Message -Level 'CRITICAL' -Component $Component -Target $Target -Data $Data
}
function Write-LogDebug {
    [CmdletBinding()]
    param([Parameter(Mandatory, Position = 0)][AllowEmptyString()][string]$Message, [string]$Component = 'Core', [AllowNull()][string]$Target, $Data)
    Write-LogEntryCore -Message $Message -Level 'DEBUG' -Component $Component -Target $Target -Data $Data
}
function Write-LogTrace {
    [CmdletBinding()]
    param([Parameter(Mandatory, Position = 0)][AllowEmptyString()][string]$Message, [string]$Component = 'Core', [AllowNull()][string]$Target, $Data)
    Write-LogEntryCore -Message $Message -Level 'TRACE' -Component $Component -Target $Target -Data $Data
}

# ---------------------------------------------------------------------------
#  Rotation and retention
# ---------------------------------------------------------------------------
function Invoke-LogRotation {
    <#
    .SYNOPSIS
        Size-based rotation: app.log -> app.log.1 -> ... -> app.log.N (oldest dropped).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)][AllowNull()][AllowEmptyString()][string]$Path,
        [double]$MaxSizeMB = $Script:LogMaxSizeMB,
        [int]$MaxFileCount = $Script:LogMaxFileCount
    )

    try {
        if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
        if (-not (Test-Path -LiteralPath $Path)) { return $false }
        $sizeMB = (Get-Item -LiteralPath $Path -Force).Length / 1MB
        if ($sizeMB -lt $MaxSizeMB) { return $false }

        $oldest = "{0}.{1}" -f $Path, $MaxFileCount
        if (Test-Path -LiteralPath $oldest) { Remove-Item -LiteralPath $oldest -Force -ErrorAction SilentlyContinue }
        for ($i = $MaxFileCount - 1; $i -ge 1; $i--) {
            $src = "{0}.{1}" -f $Path, $i
            $dst = "{0}.{1}" -f $Path, ($i + 1)
            if (Test-Path -LiteralPath $src) { Move-Item -LiteralPath $src -Destination $dst -Force -ErrorAction SilentlyContinue }
        }
        Move-Item -LiteralPath $Path -Destination ("{0}.1" -f $Path) -Force -ErrorAction Stop
        return $true
    }
    catch {
        # A failed rotation must not stop the run; the log simply keeps growing
        # until the next successful attempt.
        return $false
    }
}

function Invoke-LogRetention {
    <#
    .SYNOPSIS
        Deletes log files older than the retention window.
    .DESCRIPTION
        Any *.log or *.log.N in the resolved log directory is considered, because
        that directory is tool-owned (Logs\ under ProgramData / LocalAppData /
        script dir / TEMP). Matching by extension rather than by one hard-coded
        prefix is what makes retention work for rotating, session-stamped and
        custom -LogPath files alike. The active log file is never deleted.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory)][AllowNull()][AllowEmptyString()][string]$Path,
        [int]$RetentionDays = $Script:LogRetentionDays
    )

    $removed = 0
    try {
        if ([string]::IsNullOrWhiteSpace($Path)) { return 0 }
        if ($RetentionDays -le 0) { return 0 }
        $dir = Split-Path -Parent $Path
        if (-not $dir -or -not (Test-Path -LiteralPath $dir)) { return 0 }
        $cutoff = (Get-Date).AddDays(-$RetentionDays)

        $files = Get-ChildItem -LiteralPath $dir -File -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '\.log(\.\d+)?$' -and $_.FullName -ne $Path -and $_.LastWriteTime -lt $cutoff }
        foreach ($f in $files) {
            try {
                Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop
                $removed++
            }
            catch {
                # One undeletable old log (open handle, odd ACL) must not stop
                # the pruning of the rest.
            }
        }
    }
    catch {
        # Retention is housekeeping; a failure here is not worth aborting a run.
    }
    return $removed
}

function Write-EventLogRecord {
    <#
    .SYNOPSIS
        Optional Windows Event Log integration.
    .DESCRIPTION
        FINDING A23: the v3 README advertised event-log integration that was never
        implemented. Writes to the Application log under the PrintSpoolerCleanup
        source. Source creation needs elevation; when it fails we warn once and
        disable event logging for the run - it is an enhancement, never a
        dependency, and the file log continues unaffected.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Message,
        [ValidateSet('Information', 'Warning', 'Error', 'SuccessAudit')][string]$EntryType = 'Information',
        [int]$EventId = $Script:EventLogIdInfo
    )

    if (-not $Script:RuntimeState.EventLogEnabled) { return $false }
    if (-not (Test-RunningOnWindows)) { return $false }

    try {
        if (-not [System.Diagnostics.EventLog]::SourceExists($Script:EventLogSourceName)) {
            if (-not (Test-RunningAsAdministrator)) {
                Write-LogDebug "Event source '$($Script:EventLogSourceName)' is missing and this session is not elevated; skipping the event write." -Component 'EventLog'
                $Script:RuntimeState.EventLogEnabled = $false
                return $false
            }
            [System.Diagnostics.EventLog]::CreateEventSource($Script:EventLogSourceName, $Script:EventLogLogName)
        }
        $body = "{0}`r`nVersion : {1}`r`nLog     : {2}" -f $Message, $Script:ScriptVersion, $Script:RuntimeState.LogPath
        Write-EventLog -LogName $Script:EventLogLogName -Source $Script:EventLogSourceName -EventId $EventId -EntryType $EntryType -Message $body -Category 0 -ErrorAction Stop
        return $true
    }
    catch {
        Write-LogWarn "Event Log write failed: $($_.Exception.Message)" -Component 'EventLog'
        $Script:RuntimeState.EventLogEnabled = $false
        return $false
    }
}

function Get-LogBufferSnapshot {
    <#
    .SYNOPSIS
        Returns a copy of the in-memory log ring buffer, for diagnostics bundles.
    #>
    [CmdletBinding()]
    param()
    $state = $Script:RuntimeState
    if (-not $state.BufferSync) { return @($state.Buffer.ToArray()) }
    [System.Threading.Monitor]::Enter($state.BufferSync)
    try { return @($state.Buffer.ToArray()) }
    finally { [System.Threading.Monitor]::Exit($state.BufferSync) }
}

function Write-LogHeader {
    <#
    .SYNOPSIS
        Writes a run separator so consecutive runs in one file are easy to tell apart.
    #>
    [CmdletBinding()]
    param([hashtable]$Context = @{})

    $sep = '-' * 78
    $kv = ($Context.GetEnumerator() | Sort-Object Name | ForEach-Object { "{0}={1}" -f $_.Key, $_.Value }) -join ' '
    Write-LogEntryCore -Message $sep -Level 'INFO' -Component 'Session' -NoConsole
    Write-LogEntryCore -Message ("Run started - {0} v{1} - {2}" -f $Script:ScriptName, $Script:ScriptVersion, $kv) -Level 'INFO' -Component 'Session' -NoConsole
    Write-LogEntryCore -Message $sep -Level 'INFO' -Component 'Session' -NoConsole
}

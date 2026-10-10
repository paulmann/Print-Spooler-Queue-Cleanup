# ============================================================================
#  RUNTIME STATE  (template - token-substituted at build time)
# ----------------------------------------------------------------------------
#  One mutable hashtable holds everything the helper functions need to know
#  about the current invocation. Centralising it means:
#    * no $global: leakage (finding A21: v3 assigned $global:LASTEXITCODE)
#    * dot-sourcing the script does not terminate or pollute the host session
#    * the -Parallel / runspace workers receive plain values, not live state
# ============================================================================

function Initialize-RuntimeState {
    <#
    .SYNOPSIS
        Creates (or resets) $Script:RuntimeState for this invocation.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][string]$ScriptDirectoryOverride,
        [switch]$Reset
    )

    # ---- locale pinning: do this before anything formats a number -----------
    # Output is stated with a DOT as the decimal separator, in the documentation
    # ("1.50 KB", "184.32 MB"), in the release notes and in the test suite. Number
    # formatting follows the process culture, so on a host set to, say, ru-BY the
    # same code emitted "1,50 KB": the tool contradicted its own documented output
    # and the tests that assert it failed. Sizes and durations are machine-read
    # (reports, JSON, dashboards), so the culture is pinned here, once, for every
    # form that reaches this function: the standalone epilogue, the module header
    # and a repeat in-process call. A caller that wants its own culture back can
    # read the previous value from the runtime state's HostCultureBefore.
    $Script:HostCultureBefore = [Threading.Thread]::CurrentThread.CurrentCulture
    try {
        $pinnedCulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
        [Threading.Thread]::CurrentThread.CurrentCulture = $pinnedCulture
        [Threading.Thread]::CurrentThread.CurrentUICulture = $pinnedCulture
    }
    catch {
        # A host with trimmed culture data refuses GetCultureInfo. The tool then
        # keeps the host's culture and the helpers that promise a stable format
        # still use InvariantCulture explicitly, so nothing breaks.
    }

    # Test-Path rather than a truthiness test: reading an unset variable throws
    # under Set-StrictMode, and the build epilogue calls this function BEFORE the
    # variable has ever been assigned. A strict-mode host would therefore fail at
    # import time with "The variable '$Script:RuntimeState' cannot be retrieved".
    if (-not $Reset -and (Test-Path variable:Script:RuntimeState) -and $Script:RuntimeState) {
        return $Script:RuntimeState
    }

    # Resolve our own directory robustly: -File, dot-source, module import and
    # "pasted into a console" all report differently.
    $dir = $ScriptDirectoryOverride
    if ([string]::IsNullOrWhiteSpace($dir)) {
        if ($PSScriptRoot) { $dir = $PSScriptRoot }
        elseif ($MyInvocation.MyCommand.Path) { $dir = Split-Path -Parent $MyInvocation.MyCommand.Path }
        elseif ($PSCommandPath) { $dir = Split-Path -Parent $PSCommandPath }
        else { $dir = (Get-Location).Path }
    }

    $glyphs = $Script:GlyphAscii
    $colors = $Script:ColorMap

    $state = [ordered]@{
        # identity / paths
        ScriptDirectory      = $dir
        InstallDirectory     = $dir
        LogPath              = $null
        LogSource            = $null
        StateDirectory       = $dir

        # run bookkeeping
        RunStartedUtc        = [DateTime]::UtcNow
        RunStopwatch         = [System.Diagnostics.Stopwatch]::StartNew()
        InvocationId         = [Guid]::NewGuid().ToString('N').Substring(0, 12)
        ExitCode             = 0
        Mode                 = 'Cleanup'
        Results              = @()
        Summary              = $null

        # console rendering
        Glyphs               = $glyphs
        Colors               = $colors
        UseAscii             = $true
        NoColor              = $false
        Quiet                = $false
        ShowTimestamp        = $true
        ConsoleWidth         = 120
        NoProgress           = $false
        MachineReadable      = $false
        VerboseTranscript    = $false

        # logging
        LogLevelMap          = $Script:LogLevels
        MinLogLevelIndex     = 2
        JsonLog              = $false
        CaptureBuffer        = $true
        Buffer               = (New-Object System.Collections.Generic.List[object])
        BufferMax            = 1000
        BufferSync           = (New-Object object)
        LogSync              = (New-Object object)
        LogFailureWarned     = $false
        EventLogEnabled      = $false

        # environment facts
        IsWindows            = (Test-RunningOnWindows)
        IsAdmin              = $false
        NonInteractive       = $false
        LocalNames           = @()
        PSVersion            = $PSVersionTable.PSVersion.ToString()

        # execution
        PayloadScriptBlock   = $null
        PayloadText          = $null
        Configuration        = $null
        ConfigSource         = 'Defaults'
        ConfigPath           = $null
        ConfigProblems       = @()
        WatchdogJob          = $null
        UpdateInfo           = $null
        DotSourced           = $false
        RelaunchedElevated   = $false
        TargetList           = (New-Object System.Collections.Generic.List[string])
    }
    $Script:RuntimeState = $state

    $ctx = Get-CurrentUserContext
    $state.IsAdmin = $ctx.IsAdmin
    $state.NonInteractive = (-not $ctx.IsInteractive)
    $state.LocalNames = @(Get-LocalComputerNames)

    return $state
}

function Get-RuntimeState {
    <#
    .SYNOPSIS
        Accessor that lazily initialises state (safe for module import order).
    .NOTES
        Every read of $Script:RuntimeState is guarded by Test-Path so that a
        caller running under Set-StrictMode -Version Latest cannot throw on an
        unassigned variable. This function is the only supported way to reach the
        state object from outside the entry point.
    #>
    [CmdletBinding()]
    param()
    if (-not (Test-Path variable:Script:RuntimeState) -or -not $Script:RuntimeState) {
        $null = Initialize-RuntimeState
    }
    return $Script:RuntimeState
}

function Set-RuntimeLogLevel {
    <#
    .SYNOPSIS
        Translates a level name into the numeric threshold used by the logger.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowNull()][string]$Level)

    $state = Get-RuntimeState
    if ([string]::IsNullOrWhiteSpace($Level)) { $Level = $Script:DefaultLogLevel }
    $key = $Level.Trim().ToUpperInvariant()
    if (-not $state.LogLevelMap.Contains($key)) { $key = $Script:DefaultLogLevel }
    $state.MinLogLevelIndex = [int]$state.LogLevelMap[$key]
    return $state.MinLogLevelIndex
}

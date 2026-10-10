#region Shared constants (inlined into every build target)

# ============================================================================
#  Metadata
# ----------------------------------------------------------------------------
#  Single source of truth for version/identity. tools/Build.ps1 stamps this
#  into the .ps1, the .psd1, the .bat and the release notes, so the four can
#  never drift apart again (finding A16).
# ============================================================================
$Script:ScriptName          = 'Clear-PrintSpoolerQueue'
$Script:ProductFamily       = 'PrintSpoolerCleanup'
$Script:ScriptVersion       = '4.0.1'
$Script:ScriptBuild         = '@@BUILD_STAMP@@'
$Script:ScriptAuthor        = 'Mikhail Deynekin'
$Script:ScriptAuthorEmail   = 'Mikhail@Deynekin.com'
$Script:ScriptAuthorSite    = 'https://Deynekin.com'
$Script:ScriptRepoOwner     = 'paulmann'
$Script:ScriptRepoName      = 'Print-Spooler-Queue-Cleanup'
$Script:ScriptRepoUrl       = 'https://github.com/paulmann/Print-Spooler-Queue-Cleanup'
$Script:ScriptWikiUrl       = 'https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki'
$Script:ScriptIssuesUrl     = 'https://github.com/paulmann/Print-Spooler-Queue-Cleanup/issues'
$Script:ScriptLicense       = 'GPL-3.0-or-later'
$Script:MinPSVersion        = [Version]'5.1'
$Script:RecommendedPSVersion = [Version]'7.4'

# ============================================================================
#  Windows services
# ----------------------------------------------------------------------------
#  FINDING A1: the v3.x scripts stopped ONLY 'Spooler'. On Windows 10/11 and
#  Server the Spooler has dependents that hold open handles to spool files:
#
#    Fax                     - depends on Spooler; keeps a job channel open
#    PrintNotify             - "Printer Extensions and Notifications"; the
#                              single most common reason a Spooler stop hangs
#    PrintWorkflowUserSvc_*  - XPS/IPP "Print Workflow" per-user services
#    WSDPrintDevice*         - vendor web-service monitor helpers
#
#  v4 stops them first (in dependency order) and restarts them afterwards,
#  recording every state transition so a failure is diagnosable.
# ============================================================================
$Script:SpoolerServiceName  = 'Spooler'
$Script:DependentServices   = @(
    'PrintWorkflowUserSvc'   # prefix-matched: real names are PrintWorkflowUserSvc_1a2b3c
    'Fax'
    'PrintNotify'
)
# Services that must be running for printing but are NOT spooler dependents.
$Script:CompanionServices   = @('RpcSs', 'RpcEptMapper', 'LanmanWorkstation', 'LanmanServer')

# Spool file extensions that represent a queued print job. NEVER extend this
# list without also extending the safety guards in Test-SpoolDirectorySafe.
$Script:SpoolExtensions     = @('.SPL', '.SHD')
# Legacy/edge extensions that some drivers leave behind in PRINTERS\.
$Script:LegacySpoolExtensions = @('.SPF', '~DF')

# ============================================================================
#  Registry
# ============================================================================
$Script:RegPrintersKey        = 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers'
$Script:RegSpoolDirectoryVal  = 'DefaultSpoolDirectory'
$Script:RegMonitorsKey        = 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Monitors'
$Script:RegEnvironmentsKey    = 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments'
# Print spooler crash/hang instrumentation used by diagnostics.
$Script:RegProvidersKey       = 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Providers'

# ============================================================================
#  Default spool directory
# ----------------------------------------------------------------------------
#  Resolved at runtime from DefaultSpoolDirectory (finding: v3 honoured it but
#  never validated the value, so a corrupt REG_SZ such as "\0" or a UNC path
#  could point the delete step at an arbitrary tree). v4 validates hard.
# ============================================================================
$Script:DefaultSpoolDirectory = 'System32\spool\PRINTERS'
$Script:SpoolSubdirectories   = @('PRINTERS')   # the only leaf names we will ever delete inside

# ============================================================================
#  GitHub / update endpoints (finding A8)
# ----------------------------------------------------------------------------
#  Three sources are tried in order so that the updater still works when the
#  REST API is rate-limited (60 req/h per IP, unauthenticated) or blocked:
#    1. REST API  /releases/latest   - authoritative, has assets + checksums
#    2. REST API  /tags              - works on repos without formal releases
#    3. raw file  the source of truth - last resort; parses the version stamp
#                  out of the constants file (see UpdateSourceFilePath)
# ============================================================================
$Script:GitHubApiBase        = 'https://api.github.com/repos'
$Script:GitHubRawBase        = 'https://raw.githubusercontent.com'
$Script:UpdateDefaultBranch  = 'main'
# The one file that carries the version stamp. It MUST be a file that exists in
# the repository: the built artifact lives in dist/, which is git-ignored, so
# there is no *.ps1 script at the repository root to fetch. Read by strategy 3
# as a VERSION CHECK ONLY - it is a constants fragment, never a download source.
$Script:UpdateSourceFilePath = 'src/Shared/00-Constants.ps1'
$Script:UpdateStateFileName  = 'update-state.json'
$Script:PendingUpdateDirName = '.update'
$Script:PendingUpdateMarker  = 'pending-update.json'
$Script:ChecksumAssetName    = 'checksums.txt'
$Script:UpdateUserAgent      = 'PrintSpoolerCleanup/{0} (+{1})'
$Script:UpdateCooldownHours  = 24
$Script:UpdateTimeoutSeconds = 20

# ============================================================================
#  Logging
# ----------------------------------------------------------------------------
#  FINDING A14: v3 wrote its log *next to the script*. When installed under
#  C:\Program Files or a read-only network share every Add-Content failed and
#  the run degraded to console-only. v4 picks the first WRITABLE candidate and
#  adds rotation + retention + an optional machine-readable JSON log.
# ============================================================================
$Script:LogLevels = [ordered]@{
    TRACE    = 0
    DEBUG    = 1
    INFO     = 2
    SUCCESS  = 2   # same threshold as INFO, rendered differently
    WARN     = 3
    ERROR    = 4
    CRITICAL = 5
}
$Script:LogLevelNames       = @('TRACE', 'DEBUG', 'INFO', 'SUCCESS', 'WARN', 'ERROR', 'CRITICAL')
$Script:DefaultLogLevel     = 'INFO'
$Script:LogFileNamePrefix   = 'PrintSpoolerCleanup'
$Script:LogRetentionDays    = 30
$Script:LogMaxSizeMB        = 10
$Script:LogMaxFileCount     = 9          # PrintSpoolerCleanup.log + .1 .. .9
$Script:LogCandidateRoots   = @(
    'ProgramData',      # %ProgramData%\PrintSpoolerCleanup\Logs  (machine-wide, writable)
    'LocalAppData',     # %LocalAppData%\PrintSpoolerCleanup\Logs (per-user fallback)
    'ScriptDirectory',  # <script dir>\Logs                       (portable/green mode)
    'Temp'              # %TEMP%\PrintSpoolerCleanup\Logs         (last resort)
)
$Script:EventLogSourceName  = 'PrintSpoolerCleanup'
$Script:EventLogLogName     = 'Application'
$Script:EventLogIdInfo      = 1000
$Script:EventLogIdWarning   = 1001
$Script:EventLogIdError     = 1002
$Script:EventLogIdSuccess   = 1003
$Script:EventLogCategory    = 'Print Spooler Maintenance'

# ============================================================================
#  Configuration
# ----------------------------------------------------------------------------
#  Precedence (highest wins):
#    1. Explicit CLI parameters
#    2. Environment variables  PSQ_*
#    3. -ConfigPath <file>
#    4. Auto-discovered config file
#    5. Built-in defaults
# ============================================================================
$Script:ConfigFileName      = 'PrintSpoolerCleanup.config.psd1'
$Script:ConfigEnvPrefix     = 'PSQ_'
$Script:ConfigCandidateRoots = @(
    'ScriptDirectory',
    'CurrentDirectory',
    'ProgramData',
    'LocalAppData'
)

# ============================================================================
#  Safety: parameters for which a BLANK value is a hard error, not a no-op
# ----------------------------------------------------------------------------
#  An empty string is FALSY in PowerShell, and the cleanup payload gates every
#  selection filter on a truthiness test:
#
#      if ($Options['PrinterName'] -and @($Options['PrinterName']).Count -gt 0)
#
#  So a filter of '' does not narrow the selection - the guard fails and the run
#  silently becomes NON-SELECTIVE, cleaning EVERY printer and every job instead
#  of the one the operator named. The mirror case is worse: a blank
#  -ExcludePrinterName silently withdraws a protection that was asked for.
#
#  This is reachable in practice from the .bat launcher, where `/printer` typed
#  without a value forwards `-PrinterName ""`. Blank is therefore rejected with
#  exit code 3 (BadArguments) rather than interpreted as "match everything".
#
#  Kept here, not inline in the validator, so the list is reviewable in one
#  place and so a future selection parameter cannot be added to the product
#  without a decision about whether it belongs on this list.
# ============================================================================
$Script:BlankValueIsUnsafeParameters = @(
    # Selection filters - a blank value widens the blast radius.
    'PrinterName',
    'ExcludePrinterName',
    'OwnerName',
    'DocumentName',
    'OlderThan',
    'NewerThan',
    # Target/path parameters - a blank value means "unspecified", which for a
    # path silently falls back to a default location the operator did not choose.
    'BackupPath',
    'ArchivePath',
    'ExportPath',
    'ConfigPath',
    'RestoreDestination',
    'ComputerListFile',
    'DiagnosticsBundlePath'
)

# ============================================================================
#  Operational defaults
# ============================================================================
$Script:DefaultServiceTimeoutSeconds = 60
$Script:DefaultSettleSeconds         = 2
$Script:DefaultRetryCount            = 3
$Script:DefaultRetryDelayMs          = 500
$Script:RetryBackoffMultiplier       = 2.0   # 500ms -> 1s -> 2s (exponential)
$Script:DefaultThrottleLimit         = 8
$Script:MaxThrottleLimit             = 64
$Script:DefaultQuiescenceSeconds     = 10
$Script:DefaultWatchdogMinutes       = 30
$Script:DefaultLogRetentionDays      = 30

# ============================================================================
#  Exit codes (public contract - see docs/Exit-Codes.md, test-enforced)
# ----------------------------------------------------------------------------
#  0-9 mirror the batch launcher exactly so monitoring rules are portable.
# ============================================================================
$Script:ExitCode = [ordered]@{
    Success              = 0   # All targets cleaned, spooler verified running
    PartialSuccess       = 1   # >=1 target cleaned, >=1 target failed or warned
    Failed               = 2   # No target could be cleaned
    BadArguments         = 3   # Invalid parameter combination / unknown switch
    Cancelled            = 4   # User declined the confirmation prompt
    NotAdministrator     = 5   # Elevation required but not available
    UpdateCheckFailed    = 6   # -CheckUpdate could not reach GitHub
    UpdateAvailable      = 7   # -CheckUpdate found a newer version
    UpdateApplied        = 8   # -Update / -AutoUpdate installed a new version
    UpdateFailed         = 9   # Update download/verification/install failed
    UnsupportedPlatform  = 10  # Not Windows / Windows without a Print Spooler
    DiagnosticsOnly      = 11  # -Diagnostic completed (informational result)
    ServiceFailure       = 12  # Spooler could not be stopped or restarted
    RestoreFailed        = 13  # Restore-SpoolerBackup could not complete
    SignatureInvalid     = 14  # -VerifySignature: bad/missing/untrusted signature
    Timeout              = 15  # Watchdog or service timeout expired
    Aborted              = 130 # Terminated by Ctrl+C / host shutdown
}

# ============================================================================
#  Result status strings (public contract - appear in JSON/CSV reports)
# ============================================================================
$Script:ResultStatus = [ordered]@{
    Success         = 'Success'
    Warning         = 'SuccessWithWarnings'
    Partial         = 'PartialSuccess'
    Failed          = 'Failed'
    Skipped         = 'Skipped'
    Cancelled       = 'Cancelled'
    DryRun          = 'DryRun'
    NoWork          = 'NothingToDo'
    NotApplicable   = 'NotApplicable'
}

# ============================================================================
#  UI glyphs
# ----------------------------------------------------------------------------
#  FINDING: v3 emitted raw Unicode glyphs that rendered as garbage in the
#  default OEM code page (437/850) of cmd.exe and in redirected logs. v4 has
#  two palettes and picks ASCII automatically when the host cannot render it.
# ============================================================================
$Script:GlyphUnicode = [ordered]@{
    Info        = [char]0x2139   # information source
    Success     = [char]0x2714   # heavy check mark
    Warning     = [char]0x26A0   # warning sign
    Error       = [char]0x2718   # heavy ballot X
    Critical    = [char]0x2620   # skull and crossbones
    Progress    = [char]0x2794   # heavy wide-headed rightwards arrow
    Debug       = [char]0x2699   # gear
    Trace       = [char]0x00B7   # middle dot
    Bullet      = [char]0x2022   # bullet
    ArrowRight  = [char]0x2192   # rightwards arrow
    Box         = [char]0x25A0   # black square
    Corner      = [char]0x00BB   # right-pointing double angle quote
}
$Script:GlyphAscii = [ordered]@{
    Info        = 'i'
    Success     = '[OK]'
    Warning     = '[!]'
    Error       = '[X]'
    Critical    = '[!!]'
    Progress    = '->'
    Debug       = '[D]'
    Trace       = '[T]'
    Bullet      = '*'
    ArrowRight  = '->'
    Box         = '#'
    Corner      = '>'
}

# Colour map (ConsoleColor names are valid on every host, including ISE/VSCode).
$Script:ColorMap = [ordered]@{
    Info        = 'Gray'
    Success     = 'Green'
    Warning     = 'Yellow'
    Error       = 'Red'
    Critical    = 'DarkRed'
    Progress    = 'Cyan'
    Debug       = 'DarkGray'
    Trace       = 'DarkGray'
    Header      = 'Cyan'
    Label       = 'DarkGray'
    Value       = 'White'
    Highlight   = 'Magenta'
}

# ============================================================================
#  Safety guard-rails
# ----------------------------------------------------------------------------
#  The delete step is the only destructive operation in the product, so it is
#  fenced by four independent guards. If ANY guard fails, the run aborts with
#  ExitCode.Failed and nothing is deleted.
# ============================================================================
# Minimum number of path segments in the resolved spool directory.
#
# Calibrated deliberately, not arbitrarily: 3 accepts the OS default
# (C:\Windows\System32\spool\PRINTERS, 5 segments) and any sane relocated
# spool root (D:\Spool\PRINTERS), while rejecting 'C:\PRINTERS' and '\\host\x'.
# Going higher would break administrators who legitimately relocate the spool
# directory - the guard must stop accidents, not policy.
#
# Depth is the LAST of four layers; the primary protections are the PRINTERS
# leaf requirement, the drive-letter shape check and the forbidden-root list.
$Script:SpoolGuardMinPathDepth   = 3
# Static literals only. Machine-specific roots (%SystemRoot%, %ProgramFiles%,
# %USERPROFILE%, ...) are added at runtime by Get-SpoolGuardForbiddenRoots so
# that an unset environment variable can never inject $null into the list.
$Script:SpoolGuardForbiddenRoots = @(
    'C:\', 'C:\Windows', 'C:\Windows\System32', 'C:\Windows\SysWOW64',
    'C:\Users', 'C:\Program Files', 'C:\Program Files (x86)', 'C:\ProgramData',
    'C:\Recovery', 'C:\PerfLogs', 'C:\Boot'
)
$Script:SpoolGuardRequiredLeaf   = 'PRINTERS'   # resolved dir must end with this (case-insensitive)
$Script:SpoolGuardMaxFileCount   = 100000       # refuse to bulk-delete absurd file counts
$Script:SpoolGuardMaxTotalBytes  = 100GB        # refuse absurd total sizes (indicates wrong dir)

#endregion

# ============================================================================
#  Comprehensive built-in help (finding A9)
# ----------------------------------------------------------------------------
#  Requirement: one switch that answers every question an operator or an
#  auditor can have without leaving the console, including the parts that are
#  normally buried in a wiki - exit codes, log schema, report schema, config
#  precedence, remote prerequisites, update semantics and the safety model.
#
#  The parameter table is generated from the real parameter metadata of
#  Clear-PrintSpoolerQueue at runtime, so the help can never drift away from
#  the code (this is asserted by the test suite).
# ============================================================================

function Get-HelpSectionNames {
    <#
    .SYNOPSIS
        Canonical list of -Help sections (also used for tab-completion and tests).
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param()
    return @(
        'All', 'Usage', 'QuickStart', 'Modes', 'Parameters', 'Examples',
        'ExitCodes', 'Configuration', 'Environment', 'Logging', 'Reports',
        'Remote', 'Scheduling', 'Updates', 'Safety', 'Troubleshooting',
        'Diagnostics', 'Version', 'Credits'
    )
}

function Get-HelpText {
    <#
    .SYNOPSIS
        Renders the built-in reference documentation.
    .PARAMETER Section
        One of the names returned by Get-HelpSectionNames. 'All' renders everything.
    .PARAMETER Search
        Filters the output to blocks that contain the search term.
    .PARAMETER Width
        Wrap width. Defaults to the detected console width.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [ValidateSet('All', 'Usage', 'QuickStart', 'Modes', 'Parameters', 'Examples',
                     'ExitCodes', 'Configuration', 'Environment', 'Logging', 'Reports',
                     'Remote', 'Scheduling', 'Updates', 'Safety', 'Troubleshooting',
                     'Diagnostics', 'Version', 'Credits')]
        [string]$Section = 'All',
        [AllowNull()][string]$Search,
        [int]$Width = 0,
        [switch]$IncludeParameterMetadata
    )

    if ($Width -le 0) {
        try { $Width = [Math]::Max(80, [int]$Host.UI.RawUI.BufferSize.Width) } catch { $Width = 100 }
    }
    $Width = [Math]::Min($Width, 140)
    $nl = [Environment]::NewLine

    # Local scriptblock closures, NOT nested functions: a nested `function`
    # declaration would be re-created on every call and would leak a command name
    # into script scope once the build inlines this file.
    #
    # The names are deliberately multi-character. PowerShell variables are
    # CASE-INSENSITIVE, so closures named $W/$P/$Sub collide with the loop
    # variables $w/$p/$section used further down and get silently overwritten
    # mid-render - which presents as a baffling "'W' is not recognized" error
    # hundreds of lines away from the actual mistake.
    $sb = New-Object System.Text.StringBuilder
    $WriteLine = { param([string]$t = '') [void]$sb.AppendLine($t) }
    $WriteRule = { param([string]$c = '=') [void]$sb.AppendLine(($c * $Width)) }
    $WriteHead = { param([string]$t) [void]$sb.AppendLine(''); & $WriteRule '='; [void]$sb.AppendLine("  $t"); & $WriteRule '=' }
    $WriteSub  = { param([string]$t) [void]$sb.AppendLine(''); $rule = '-' * [Math]::Max(3, $Width - $t.Length - 9); [void]$sb.AppendLine('  --- ' + $t + ' ' + $rule) }
    $WritePara = { param([string]$t) foreach ($line in (Split-HelpParagraph -Text $t -Width ($Width - 4))) { [void]$sb.AppendLine("    $line") } }

    $sections = if ($Section -eq 'All') { Get-HelpSectionNames | Where-Object { $_ -ne 'All' } } else { @($Section) }

    # ------------------------------------------------------------------ header
    & $WriteLine ("{0} v{1}" -f 'Print Spooler Queue Cleanup', $Script:ScriptVersion)
    & $WriteLine ("{0}  |  {1}  |  PowerShell {2}+" -f $Script:ScriptRepoUrl, $Script:ScriptLicense, $Script:MinPSVersion)
    & $WriteLine ''

    foreach ($sec in $sections) {
        switch ($sec) {

'Usage' {
    & $WriteHead 'USAGE'
    & $WriteLine '  Clear-PrintSpoolerQueue.ps1 [-ComputerName <string[]>] [-Force] [-WhatIf]'
    & $WriteLine '                              [-PrinterName <string[]>] [-OlderThan <string>]'
    & $WriteLine '                              [-Backup] [-Diagnostic] [-Repair]'
    & $WriteLine '                              [-CheckUpdate] [-Update] [-Help] [-HelpSection <name>]'
    & $WriteLine '                              [<CommonParameters>]'
    & $WriteLine ''
    & $WriteLine '  Clear-PrintSpoolerQueue.bat [/Force] [/WhatIf] [/Diagnostic] [/Repair]'
    & $WriteLine '                              [/Update] [/CheckUpdate] [/?] [<options>]'
    & $WriteLine ''
    & $WritePara 'The .bat file is a thin, self-elevating launcher. It detects PowerShell 7 (pwsh.exe) and falls back to Windows PowerShell 5.1, applies any staged update, then hands over to the .ps1 with the equivalent arguments. For automation, call the .ps1 directly.'
    & $WriteLine ''
    & $WriteLine '  Import as a module instead of calling the script:'
    & $WriteLine '    Import-Module .\PrintSpoolerCleanup\PrintSpoolerCleanup.psd1'
    & $WriteLine '    Clear-PrintSpoolerQueue -ComputerName SRV01 -Force -Backup'
}

'QuickStart' {
    & $WriteHead 'QUICK START'
    & $WriteSub 'First contact (safe, changes nothing)'
    & $WriteLine '    .\Clear-PrintSpoolerQueue.ps1 -WhatIf'
    & $WriteLine '    .\Clear-PrintSpoolerQueue.ps1 -Diagnostic'
    & $WriteLine ''
    & $WriteSub 'Fix one machine right now'
    & $WriteLine '    .\Clear-PrintSpoolerQueue.ps1 -Force -Backup'
    & $WriteLine ''
    & $WriteSub 'Fix a whole site'
    & $WriteLine '    .\Clear-PrintSpoolerQueue.ps1 -ComputerName (Get-Content .\workstations.txt) -Force -Backup -ThrottleLimit 16'
    & $WriteLine ''
    & $WriteSub 'Only the jobs that are actually stuck'
    & $WriteLine '    .\Clear-PrintSpoolerQueue.ps1 -OlderThan 30m -Force -Backup'
    & $WriteLine '    .\Clear-PrintSpoolerQueue.ps1 -PrinterName "HP*","Canon*" -Force'
    & $WriteLine ''
    & $WriteSub 'Deep clean a machine that keeps re-sticking'
    & $WriteLine '    .\Clear-PrintSpoolerQueue.ps1 -Repair -Force -Backup'
    & $WriteLine ''
    & $WriteSub 'Schedule nightly maintenance'
    & $WriteLine '    Install-SpoolerCleanupTask -ScriptPath C:\Tools\Clear-PrintSpoolerQueue.ps1 -Schedule Daily -StartTime 03:00'
    & $WriteLine ''
    & $WriteSub 'Keep the tool current'
    & $WriteLine '    .\Clear-PrintSpoolerQueue.ps1 -CheckUpdate'
    & $WriteLine '    .\Clear-PrintSpoolerQueue.ps1 -Update -Force'
}

'Modes' {
    & $WriteHead 'OPERATING MODES'
    & $WritePara 'Exactly one mode is active per run. Modes are mutually exclusive by design so a scheduled task can never accidentally become a destructive one.'
    & $WriteLine ''
    $modes = @(
        @{ M = 'Cleanup (default)'; D = 'Stop Spooler -> delete selected spool files -> restart Spooler -> verify.' }
        @{ M = '-WhatIf';           D = 'Full inventory and plan, zero changes. The service is NOT stopped.' }
        @{ M = '-Diagnostic';       D = 'Read-only deep inspection: services, processes, printers, drivers, ports, spool directory, job table, event logs.' }
        @{ M = '-Repair';           D = 'Cleanup PLUS root-cause repair: dependent-service cycling, PrintNotify recovery, Forms cache purge, optional driver-store purge, optional spool-directory reset.' }
        @{ M = '-Restore';          D = 'Forensic restore of a backup bundle created by -Backup.' }
        @{ M = '-CheckUpdate';      D = 'Query GitHub for a newer release and exit (exit code 7 when an update exists).' }
        @{ M = '-Update';           D = 'Download, verify (SHA-256), stage and optionally apply a newer version.' }
        @{ M = '-SelfTest';         D = 'Run the internal consistency checks (payload parse, guards, schema) and exit.' }
        @{ M = '-Help';             D = 'This document, in full.' }
        @{ M = '-HelpSection <name>'; D = 'One section of this document (-ListHelpSections lists the names).' }
    )
    foreach ($m in $modes) { & $WriteLine ("  {0,-22} {1}" -f $m.M, $m.D) }
}

'Parameters' {
    & $WriteHead 'PARAMETERS'
    $meta = Get-ParameterDocumentation
    if ($IncludeParameterMetadata) {
        & $WriteLine ''
        & $WriteLine '  (generated from live parameter metadata - cannot drift from the code)'
    }
    $currentGroup = $null
    foreach ($p in $meta) {
        if ($p.Group -ne $currentGroup) { & $WriteSub $p.Group; $currentGroup = $p.Group }
        $type = $p.Type
        $def  = if ($p.Default) { "  [default: $($p.Default)]" } else { '' }
        & $WriteLine ("  -{0} <{1}>{2}" -f $p.Name, $type, $def)
        foreach ($line in (Split-HelpParagraph -Text $p.Description -Width ($Width - 6))) { & $WriteLine ("      $line") }
        if ($p.Note) { foreach ($line in (Split-HelpParagraph -Text ("Note: " + $p.Note) -Width ($Width - 6))) { & $WriteLine ("      $line") } }
        & $WriteLine ''
    }
}

'Examples' {
    & $WriteHead 'EXAMPLES'
    $examples = @(
        @{ T = 'Inspect before touching anything'; C = @'
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic -ExportFormat Html -ExportPath C:\Temp
'@; N = 'Writes a self-contained HTML report you can attach to the ticket.' }
        @{ T = 'Clear the local queue with a backup'; C = @'
.\Clear-PrintSpoolerQueue.ps1 -Force -Backup
'@; N = 'Backup lands in %ProgramData%\PrintSpoolerCleanup\Backups by default.' }
        @{ T = 'Delete only jobs older than two hours'; C = @'
.\Clear-PrintSpoolerQueue.ps1 -OlderThan 2h -Force -Backup
'@; N = 'Jobs submitted in the last two hours keep printing. Ideal for a busy shared queue.' }
        @{ T = 'Clear one printer on one server'; C = @'
.\Clear-PrintSpoolerQueue.ps1 -ComputerName SRV-PRINT01 -PrinterName "KYOCERA TASKalfa*" -Force
'@; N = 'Wildcards are supported. -ExcludePrinterName protects the rest.' }
        @{ T = 'Delete specific job ids'; C = @'
.\Clear-PrintSpoolerQueue.ps1 -JobId 4217,4218 -Force -Backup
'@ }
        @{ T = 'Fleet-wide maintenance in parallel'; C = @'
$targets = Get-ADComputer -Filter {OperatingSystem -like "Windows*"} -SearchBase "OU=Workstations,DC=corp,DC=local" | Select-Object -ExpandProperty Name
.\Clear-PrintSpoolerQueue.ps1 -ComputerName $targets -Credential $cred -ThrottleLimit 24 -Force -Backup -ExportFormat Csv -ExportPath \\filesrv\logs\spooler
'@; N = 'Requires PowerShell 7 for -ThrottleLimit > 1. On 5.1 a runspace pool is used automatically.' }
        @{ T = 'Pipe computers from any source'; C = @'
Get-Content .\hosts.txt | .\Clear-PrintSpoolerQueue.ps1 -Force -Backup
Import-Csv .\assets.csv | .\Clear-PrintSpoolerQueue.ps1 -Force   # binds on a ComputerName column
'@ }
        @{ T = 'Deep repair a machine that keeps re-sticking'; C = @'
.\Clear-PrintSpoolerQueue.ps1 -Repair -PurgeFormsCache -Force -Backup -Verbose
'@ }
        @{ T = 'Remove unused printer drivers (destructive - needs -Force)'; C = @'
.\Clear-PrintSpoolerQueue.ps1 -Repair -PurgeDriverStore -Force -Backup
'@; N = 'Only driver packages with no printer referencing them are removed.' }
        @{ T = 'Scheduled, unattended, machine-readable'; C = @'
.\Clear-PrintSpoolerQueue.ps1 -Force -NonInteractive -Quiet -NoProgress -MachineReadableJson -OlderThan 6h -Backup
'@; N = 'stdout is then a single JSON document - perfect for RMM ingestion.' }
        @{ T = 'Restore a backup for forensic analysis'; C = @'
Restore-SpoolerBackup -Path "C:\ProgramData\PrintSpoolerCleanup\Backups\spool-backup_SRV01_20261010-031500.zip" -DestinationDirectory C:\Temp\restored
'@; N = 'Copying spool files back into PRINTERS does NOT re-queue them. Re-submit from the application.' }
        @{ T = 'Check for a new version without changing anything'; C = @'
.\Clear-PrintSpoolerQueue.ps1 -CheckUpdate
if ($LASTEXITCODE -eq 7) { .\Clear-PrintSpoolerQueue.ps1 -Update -Force }
'@ }
        @{ T = 'Health probe for monitoring'; C = @'
Test-PrintSpoolerHealth -ComputerName SRV-PRINT01 | ConvertTo-Json
'@; N = 'Exit code 0 = healthy, 1 = degraded, 2 = unhealthy.' }
        @{ T = 'Install nightly maintenance'; C = @'
Install-SpoolerCleanupTask -ScriptPath C:\Tools\Clear-PrintSpoolerQueue.ps1 -Schedule Daily -StartTime 03:00 -ComputerListFile C:\Tools\hosts.txt
Get-SpoolerCleanupTask
'@ }
        @{ T = 'Package a bug report'; C = @'
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic -DiagnosticsBundlePath C:\Temp
'@; N = 'Produces a ZIP with environment, diagnostics, effective config (secrets redacted), log buffer and file hashes.' }
    )
    $i = 0
    foreach ($e in $examples) {
        $i++
        & $WriteLine ''
        & $WriteLine ("  EXAMPLE $i - {0}" -f $e['T'])
        foreach ($line in ($e['C'] -split "\r?\n")) { if ($line.Trim()) { & $WriteLine ("    $line") } }
        # Index access, NOT $e.N. The optional note key 'N' is absent from some
        # entries, and reading a missing hashtable key as a property throws
        # "The property 'N' cannot be found" under Set-StrictMode -Version
        # Latest. The product's own handler routes that exception to the log
        # and exits 2, so on a hardened host the ENTIRE reference rendered as
        # silent empty output (defect 48 - the same class as defect 29).
        if ($e['N']) { & $WriteLine ''; & $WritePara $e['N'] }
    }
}

'ExitCodes' {
    & $WriteHead 'EXIT CODES'
    & $WritePara 'Identical for the .ps1 and the .bat, so a monitoring rule written for one works for the other. The value is available as $LASTEXITCODE after the process exits.'
    & $WriteLine ''
    $codes = @(
        @{ C = 0;   N = 'Success';              D = 'Every target processed, Spooler verified Running.' }
        @{ C = 1;   N = 'PartialSuccess';       D = 'Some targets succeeded, some failed or produced warnings.' }
        @{ C = 2;   N = 'Failed';               D = 'No target could be processed.' }
        @{ C = 3;   N = 'BadArguments';         D = 'Invalid parameter combination or unreadable input.' }
        @{ C = 4;   N = 'Cancelled';            D = 'The operator declined the confirmation prompt.' }
        @{ C = 5;   N = 'NotAdministrator';     D = 'Elevation required and not available.' }
        @{ C = 6;   N = 'UpdateCheckFailed';    D = '-CheckUpdate could not reach GitHub (offline, proxy, rate limit).' }
        @{ C = 7;   N = 'UpdateAvailable';      D = '-CheckUpdate found a newer version. Nothing was changed.' }
        @{ C = 8;   N = 'UpdateApplied';        D = '-Update/-AutoUpdate installed a newer version.' }
        @{ C = 9;   N = 'UpdateFailed';         D = 'Download, checksum verification or staging failed.' }
        @{ C = 10;  N = 'UnsupportedPlatform';  D = 'Not Windows, or the Print Spooler feature is not installed.' }
        @{ C = 11;  N = 'DiagnosticsOnly';      D = '-Diagnostic completed. Informational, not an error.' }
        @{ C = 12;  N = 'ServiceFailure';       D = 'The Spooler could not be stopped or restarted.' }
        @{ C = 13;  N = 'RestoreFailed';        D = 'Restore-SpoolerBackup could not complete.' }
        @{ C = 14;  N = 'SignatureInvalid';     D = '-VerifySignature found a broken, missing or unexpected signature.' }
        @{ C = 15;  N = 'Timeout';              D = 'The watchdog expired and the run was terminated.' }
        @{ C = 130; N = 'Aborted';              D = 'Interrupted with Ctrl+C or by host shutdown.' }
    )
    foreach ($c in $codes) { & $WriteLine ("  {0,4}  {1,-20} {2}" -f $c.C, $c.N, $c.D) }
    & $WriteLine ''
    & $WritePara 'Monitoring tip: alert on 2, 5, 12 and 15. Treat 1 as "needs attention", 7 as "plan an update", and 0/11 as healthy.'
}

'Configuration' {
    & $WriteHead 'CONFIGURATION'
    & $WriteSub 'Precedence (highest wins)'
    & $WriteLine '    1. Command-line parameters'
    & $WriteLine '    2. PSQ_* environment variables'
    & $WriteLine '    3. -ConfigPath <file>'
    & $WriteLine '    4. Auto-discovered PrintSpoolerCleanup.config.psd1'
    & $WriteLine '         (script dir -> current dir -> %ProgramData%\PrintSpoolerCleanup -> %LocalAppData%\PrintSpoolerCleanup)'
    & $WriteLine '    5. Built-in defaults'
    & $WriteLine ''
    & $WriteSub 'Security of the config file'
    & $WritePara 'The file is parsed with the PowerShell AST and evaluated with SafeGetValue(). Command invocations, pipelines and assignments are rejected outright, so a config that arrives from a file share or a Git repository cannot execute code. Files larger than 1 MB are refused.'
    & $WriteLine ''
    & $WriteSub 'Generate a starting point'
    & $WriteLine '    .\Clear-PrintSpoolerQueue.ps1 -SaveConfig -ConfigPath C:\Tools\PrintSpoolerCleanup.config.psd1'
    & $WriteLine ''
    & $WriteSub 'Schema'
    $defaults = Get-DefaultConfiguration
    # $configSection, NOT $section: PowerShell variables are case-insensitive, so
    # a loop variable named $section collides with this function's own
    # [ValidateSet()] $Section parameter and assigning 'Cleanup' to it throws.
    foreach ($configSection in ($defaults.Keys | Sort-Object)) {
        & $WriteLine ''
        & $WriteLine ("    $configSection = @{")
        foreach ($key in ($defaults[$configSection].Keys | Sort-Object)) {
            $v = $defaults[$configSection][$key]
            $rendered = if ($null -eq $v) { '$null' }
                        elseif ($v -is [bool]) { if ($v) { '$true' } else { '$false' } }
                        elseif ($v -is [array]) { if ($v.Count -eq 0) { '@()' } else { '@(' + (($v | ForEach-Object { "'$_'" }) -join ', ') + ')' } }
                        else { "'$v'" }
            & $WriteLine ("        {0,-28} = {1}" -f $key, $rendered)
        }
        & $WriteLine '    }'
    }
    & $WriteLine ''
    & $WriteSub 'Validation'
    & $WritePara 'Every value is range-checked and unknown keys/sections are reported as errors (exit code 3). A configuration problem is never silently ignored, because a silent default is how a site loses its backups.'
}

'Environment' {
    & $WriteHead 'ENVIRONMENT VARIABLES'
    & $WritePara 'All optional. Useful for machine-wide policy set through Group Policy preferences, or for CI.'
    & $WriteLine ''
    $vars = @(
        @{ V = 'PSQ_COMPUTERS';            D = 'Comma- or semicolon-separated target list (same as -ComputerName).' }
        @{ V = 'PSQ_FORCE';                D = 'true/1/yes -> same as -Force.' }
        @{ V = 'PSQ_LOG_PATH';             D = 'Explicit log file path.' }
        @{ V = 'PSQ_LOG_LEVEL';            D = 'TRACE|DEBUG|INFO|WARN|ERROR|CRITICAL.' }
        @{ V = 'PSQ_LOG_RETENTION_DAYS';   D = 'Rotated-log retention window.' }
        @{ V = 'PSQ_LOG_JSON';             D = 'true -> JSON Lines log for shippers.' }
        @{ V = 'PSQ_EVENT_LOG';            D = 'true -> also write to the Windows Event Log.' }
        @{ V = 'PSQ_CONFIG_PATH';          D = 'Configuration file to load.' }
        @{ V = 'PSQ_SERVICE_TIMEOUT';      D = 'Seconds to wait for a service state change.' }
        @{ V = 'PSQ_SETTLE_SECONDS';       D = 'Pause after the service restarts, before verifying.' }
        @{ V = 'PSQ_RETRY_COUNT';          D = 'Delete retry attempts per locked file.' }
        @{ V = 'PSQ_THROTTLE_LIMIT';       D = 'Parallel remote workers.' }
        @{ V = 'PSQ_BACKUP_PATH';          D = 'Backup root; implies -Backup.' }
        @{ V = 'PSQ_PRINTER_NAME';         D = 'Default -PrinterName filter list.' }
        @{ V = 'PSQ_OLDER_THAN';           D = 'Default -OlderThan filter (e.g. 2h).' }
        @{ V = 'PSQ_EXPORT_PATH';          D = 'Report export directory or file.' }
        @{ V = 'PSQ_EXPORT_FORMAT';        D = 'Comma-separated: Json,Csv,Html,Text.' }
        @{ V = 'PSQ_QUIET';                D = 'true -> suppress console decoration.' }
        @{ V = 'PSQ_NONINTERACTIVE';       D = 'true -> never prompt; refuse instead of waiting (set this for scheduled tasks and CI).' }
        @{ V = 'PSQ_MACHINE_JSON';         D = 'true -> exactly one JSON document on stdout, nothing else.' }
        @{ V = 'PSQ_NO_COLOR';             D = 'true -> plain output (also honours NO_COLOR).' }
        @{ V = 'PSQ_ASCII';                D = 'true -> ASCII glyphs for legacy code pages.' }
        @{ V = 'PSQ_UPDATE_CHANNEL';       D = 'latest | prerelease - which release to compare against.' }
        @{ V = 'PSQ_NO_UPDATE_CHECK';      D = 'true -> never check GitHub automatically.' }
        @{ V = 'PSQ_UPDATE_COOLDOWN_HOURS';D = 'Minimum hours between automatic checks (0 = always).' }
        @{ V = 'PSQ_GITHUB_TOKEN';         D = 'Raises the anonymous 60 req/h API limit to 5000.' }
    )
    foreach ($v in $vars) { & $WriteLine ("  {0,-26} {1}" -f $v.V, $v.D) }
    & $WriteLine ''
    & $WritePara 'PSQ_GITHUB_TOKEN is read from the environment and is never written to the log file or to any report.'
}

'Logging' {
    & $WriteHead 'LOGGING'
    & $WriteSub 'Where the log goes'
    & $WritePara 'The first writable location wins: -LogPath, then %ProgramData%\PrintSpoolerCleanup\Logs, then %LocalAppData%\..., then <script dir>\Logs, then %TEMP%\.... The resolved path is printed in the summary and recorded in the report. A log location is probed with a real write, never by guessing from ACLs.'
    & $WriteLine ''
    & $WriteSub 'Record format (default)'
    & $WriteLine '    2026-10-10T03:15:04.812Z|INFO    |Spooler|SRV01|Print Spooler stopped successfully'
    & $WriteLine '    <UTC ISO-8601>|<LEVEL padded to 8>|<component>|<target or ->|<message>'
    & $WriteLine ''
    & $WritePara 'Pipe-delimited with no locale-dependent formatting, so it parses identically on a German, Japanese or US host and imports cleanly into Splunk/ELK/Log Analytics.'
    & $WriteLine ''
    & $WriteSub 'JSON Lines mode'
    & $WriteLine '    .\Clear-PrintSpoolerQueue.ps1 -JsonLog'
    & $WriteLine '    {"ts":"2026-10-10T03:15:04.812Z","level":"INFO","component":"Spooler","target":"SRV01","msg":"..."}'
    & $WriteLine ''
    & $WriteSub 'Rotation and retention'
    & $WriteLine '    Rotate at 10 MB (Logging.MaxSizeMB), keep 9 generations (.1 .. .9),'
    & $WriteLine '    delete rotated files older than 30 days (Logging.RetentionDays).'
    & $WriteLine ''
    & $WriteSub 'Event Log'
    & $WriteLine '    .\Clear-PrintSpoolerQueue.ps1 -EventLog'
    & $WritePara 'Creates the source PrintSpoolerCleanup in the Application log on first use (needs elevation). IDs: 1000 info, 1001 warning, 1002 error, 1003 success. If the source cannot be created, event logging is disabled for the run and the file log continues unaffected.'
    & $WriteLine ''
    & $WriteSub 'Levels'
    & $WriteLine '    TRACE < DEBUG < INFO (= SUCCESS) < WARN < ERROR < CRITICAL'
}

'Reports' {
    & $WriteHead 'REPORTS AND MACHINE-READABLE OUTPUT'
    & $WriteSub 'Export formats'
    & $WriteLine '    -ExportFormat Json,Csv,Html,Text -ExportPath C:\Reports'
    & $WriteLine ''
    & $WritePara 'A directory produces auto-named files (Clear-PrintSpoolerQueue_<stamp>.<ext>). A file path is used for the first format. The HTML report is fully self-contained: inline CSS, no CDN, no JavaScript - safe to mail out of an air-gapped network.'
    & $WriteLine ''
    & $WriteSub 'stdout JSON for RMM/monitoring'
    & $WriteLine '    .\Clear-PrintSpoolerQueue.ps1 -Force -MachineReadableJson -Quiet'
    & $WritePara 'stdout then contains exactly one JSON document and nothing else - no banner, no progress, no colour codes. All human output is redirected to the log file.'
    & $WriteLine ''
    & $WriteSub 'Top-level schema'
    & $WriteLine '    {'
    & $WriteLine '      "schemaVersion": "4.0",'
    & $WriteLine '      "summary": {'
    & $WriteLine '        "tool","toolVersion","startedUtc","finishedUtc","durationMs","overallStatus",'
    & $WriteLine '        "mode","operator","hostComputer","powershellVersion","logPath","configSource",'
    & $WriteLine '        "targetsAttempted","targetsSucceeded","targetsWithWarnings","targetsFailed",'
    & $WriteLine '        "jobsFound","jobsSelected","jobsRemoved","filesFound","filesDeleted",'
    & $WriteLine '        "filesFailed","filesLocked","resurrectedFiles","bytesFound","bytesDeleted",'
    & $WriteLine '        "warnings","errors","backupsCreated","exitCode","results": [ ... ]'
    & $WriteLine '      }'
    & $WriteLine '    }'
    & $WriteLine ''
    & $WriteSub 'Per-computer result schema'
    & $WriteLine '    computerName,status,startedUtc,finishedUtc,durationMs,isLocal,spoolDirectory,'
    & $WriteLine '    spoolDirectorySource,serviceWasRunning,serviceIsRunning,serviceStopMs,serviceStartMs,'
    & $WriteLine '    dependentServices,jobsFound,jobsSelected,jobsRemoved,filesFound,filesSelected,'
    & $WriteLine '    filesDeleted,filesFailed,filesLocked,bytesFound,bytesDeleted,filesAfter,'
    & $WriteLine '    resurrectedFiles,backupPath,backupHash,dryRun,repairPerformed,diagnosticsCollected,'
    & $WriteLine '    warnings[],errors[],jobs[],removedFiles[],failedFiles[],orphanedFiles[],serviceTrace[]'
    & $WriteLine ''
    & $WriteSub '-PassThru'
    & $WritePara 'Returns the same object graph into the PowerShell pipeline instead of only rendering it, so you can post-process in-process:'
    & $WriteLine '    $run = .\Clear-PrintSppoolerQueue.ps1 -Force -PassThru'
    & $WriteLine '    $run.Results | Where-Object Status -ne "Success" | Format-Table ComputerName, Status'
}

'Remote' {
    & $WriteHead 'REMOTE MANAGEMENT'
    & $WriteSub 'Transport'
    & $WritePara 'WinRM over Invoke-Command. This is the only transport that works identically on Windows PowerShell 5.1 and PowerShell 7 - the -ComputerName parameter of Get-Service/Stop-Service was removed in PowerShell 6, which is exactly the bug v3.0 had.'
    & $WriteLine ''
    & $WriteSub 'Prerequisites on the target'
    & $WriteLine '    1. WinRM running:      winrm quickconfig -quiet'
    & $WriteLine '    2. Firewall open:      TCP 5985 (HTTP) or 5986 (HTTPS, with -UseSsl)'
    & $WriteLine '    3. Caller authorised:  local Administrators on the target, or -Credential'
    & $WriteLine '    4. Trusted hosts:      for workgroup machines, add the caller to'
    & $WriteLine '                           TrustedHosts on BOTH sides, or use Kerberos in a domain'
    & $WriteLine ''
    & $WriteSub 'Enable from one place'
    & $WriteLine '    Enable-PSRemoting -Force                     # on each target'
    & $WriteLine '    Set-Item WSMan:\localhost\Client\TrustedHosts -Value "10.0.0.*" -Force'
    & $WriteLine ''
    & $WriteSub 'Pre-flight probe'
    & $WritePara 'With -PreFlightProbe (default on) each target is probed with a trivial Invoke-Command and a short timeout before any real work starts. Offline hosts are reported in seconds instead of costing a full 40-second WinRM timeout each, and they never delay the healthy hosts.'
    & $WriteLine ''
    & $WriteSub 'Parallelism'
    & $WriteLine '    PowerShell 7.x  -> ForEach-Object -Parallel -ThrottleLimit N'
    & $WriteLine '    PowerShell 5.1  -> System.Management.Automation.Runspaces pool (same throttling)'
    & $WriteLine ''
    & $WritePara 'One failing computer can never abort the fleet: each worker is isolated and always produces a result record. Default throttle is 1 (sequential) because a spooler restart is disruptive; raise it deliberately.'
    & $WriteLine ''
    & $WriteSub 'Credentials'
    & $WriteLine '    $cred = Get-Credential'
    & $WriteLine '    .\Clear-PrintSpoolerQueue.ps1 -ComputerName SRV01 -Credential $cred -Force'
    & $WritePara 'Credentials are held in a PSCredential for the life of the process only. They are never written to the log file, the report, or the diagnostics bundle.'
}

'Scheduling' {
    & $WriteHead 'SCHEDULING AND AUTOMATION'
    & $WriteSub 'Install the maintenance task'
    & $WriteLine '    Install-SpoolerCleanupTask -ScriptPath C:\Tools\Clear-PrintSpoolerQueue.ps1 `'
    & $WriteLine '        -Schedule Daily -StartTime 03:00 -Backup -PassThru'
    & $WriteLine ''
    & $WriteSub 'Schedules'
    & $WriteLine '    Daily | Weekly (Mon-Fri) | Hourly (-IntervalHours N) | AtStartup | OnEvent | Once'
    & $WriteLine ''
    & $WriteSub 'Event-driven cleanup'
    & $WriteLine '    Install-SpoolerCleanupTask -ScriptPath C:\Tools\Clear-PrintSpoolerQueue.ps1 `'
    & $WriteLine '        -Schedule OnEvent -EventLogName System -EventSource "Service Control Manager" -EventId 7031'
    & $WritePara '7031 is "service terminated unexpectedly" - the exact event a crashed Spooler produces. The task then self-heals within a minute of the crash instead of waiting for a user ticket.'
    & $WriteLine ''
    & $WriteSub 'What the task actually runs'
    & $WriteLine '    <pwsh|powershell>.exe -NoProfile -ExecutionPolicy Bypass -File "<script>"'
    & $WriteLine '        -Force -NonInteractive -Quiet -NoProgress -Backup -AutoUpdate'
    & $WritePara 'Always the .ps1, never the .bat, and always with -NonInteractive so it can never block on a prompt. This is the failure mode of the schtasks one-liner that v3 documented.'
    & $WriteLine ''
    & $WriteSub 'Manage'
    & $WriteLine '    Get-SpoolerCleanupTask'
    & $WriteLine '    Start-SpoolerCleanupTask          # run now, wait for the result'
    & $WriteLine '    Uninstall-SpoolerCleanupTask'
    & $WriteLine ''
    & $WriteSub 'Manual schtasks equivalent (for GPO/ConfigMgr packaging)'
    & $WriteLine '    schtasks /Create /TN "PrintSpoolerCleanup" /SC DAILY /ST 03:00 /RU SYSTEM `'
    & $WriteLine '      /TR "powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Tools\Clear-PrintSpoolerQueue.ps1 -Force -NonInteractive -Quiet -Backup"'
}

'Updates' {
    & $WriteHead 'UPDATES'
    & $WriteSub 'Check'
    & $WriteLine '    .\Clear-PrintSpoolerQueue.ps1 -CheckUpdate'
    & $WriteLine '    Get-SpoolerCleanupUpdate -Channel stable'
    & $WritePara 'Exit code 7 means "a newer version exists"; 6 means "GitHub could not be reached". Both are safe to poll from monitoring.'
    & $WriteLine ''
    & $WriteSub 'Discovery order (three independent strategies)'
    & $WriteLine '    1. GET /repos/paulmann/Print-Spooler-Queue-Cleanup/releases/latest'
    & $WriteLine '    2. GET /repos/.../tags                (repos without formal releases)'
    & $WriteLine '    3. GET raw.githubusercontent.com/.../main/Clear-PrintSpoolerQueue.ps1'
    & $WriteLine '       and read the version stamp with the PowerShell AST'
    & $WritePara 'If the REST API is rate-limited (60 requests/hour per source IP, unauthenticated) the tool still finds the current version through strategy 3. Supply PSQ_GITHUB_TOKEN or -GitHubToken to raise the limit to 5000/hour.'
    & $WriteLine ''
    & $WriteSub 'Apply'
    & $WriteLine '    .\Clear-PrintSpoolerQueue.ps1 -Update -Force'
    & $WriteLine '    Update-SpoolerCleanup -Force -ApplyNow'
    & $WriteLine ''
    & $WriteSub 'Integrity'
    & $WritePara 'The SHA-256 published in the release asset checksums.txt is verified against the downloaded file. On mismatch the download is discarded and the installed copy is left untouched. When a release publishes no checksums the update is refused unless -SkipChecksumVerification is passed explicitly - a silent "install whatever arrived" path is how tooling becomes an attack vector.'
    & $WriteLine ''
    & $WriteSub 'Staged self-update (never overwrites a running script)'
    & $WriteLine '    <install>\.update\Clear-PrintSpoolerQueue.ps1'
    & $WriteLine '    <install>\.update\pending-update.json'
    & $WritePara 'A process cannot safely replace the file it is executing from. Updates therefore land in .update\ and are moved into place at the START of the next run (both the .bat launcher and the .ps1 do this). The previous version is kept in <install>\backup-<version>-<stamp>\ for instant rollback.'
    & $WriteLine ''
    & $WriteSub 'Automatic, quiet, rate-limited'
    & $WriteLine '    .\Clear-PrintSpoolerQueue.ps1 -AutoUpdate -Force'
    & $WritePara 'At most one check per cooldown window (default 24 h, state in update-state.json). A failed or slow check never delays or aborts the cleanup. Set PSQ_NO_UPDATE_CHECK=true in air-gapped environments.'
    & $WriteLine ''
    & $WriteSub 'Channels'
    & $WriteLine '    latest     newest non-draft release (default)'
    & $WriteLine '    stable     newest release that is not marked prerelease'
    & $WriteLine '    prerelease newest release including prereleases'
}

'Safety' {
    & $WriteHead 'SAFETY MODEL'
    & $WritePara 'This tool deletes files as an administrator on production print servers. Everything below exists because the alternative is a self-inflicted outage.'
    & $WriteLine ''
    & $WriteSub 'The delete step is fenced by four independent guards'
    & $WriteLine '    1. Extension allow-list   only .SPL/.SHD (optionally .SPF/~DF with'
    & $WriteLine '                              -PurgeLegacyExtensions). Never a wildcard.'
    & $WriteLine '    2. Path validation        the resolved spool directory must end in'
    & $WriteLine '                              PRINTERS, must be rooted, must be at least'
    & $WriteLine '                              3 segments deep, must not be a UNC path, and'
    & $WriteLine '                              must not equal a protected system root.'
    & $WriteLine '    3. Volume guards          more than 100,000 files or more than 100 GB'
    & $WriteLine '                              in the directory means "this is not a spool'
    & $WriteLine '                              directory" - the run aborts.'
    & $WriteLine '    4. No recursion           subdirectories are never enumerated for'
    & $WriteLine '                              deletion.'
    & $WritePara 'If any guard fails, nothing is deleted and the exit code is 2 (Failed). The guards apply identically on local and remote targets because there is only one implementation.'
    & $WriteLine ''
    & $WriteSub 'Registry override is validated, not trusted'
    & $WritePara 'DefaultSpoolDirectory can be set to anything, including a path that would make a naive cleanup catastrophic. A value that is a UNC path, shorter than 4 characters, or contains control characters is ignored with a warning and the OS default is used instead.'
    & $WriteLine ''
    & $WriteSub 'Service handling'
    & $WriteLine '    * Dependent services (Fax, PrintNotify, PrintWorkflowUserSvc_*) are'
    & $WriteLine '      stopped BEFORE the Spooler and restarted AFTER it.'
    & $WriteLine '    * Only services that were actually Running are restarted, so a disabled'
    & $WriteLine '      Fax service is never silently enabled by a cleanup.'
    & $WriteLine '    * A service stuck in STOP_PENDING can be escalated to a process kill'
    & $WriteLine '      (ForceStopFallback) instead of hanging the run.'
    & $WriteLine '    * If the Spooler cannot be stopped, NOTHING is deleted and the service'
    & $WriteLine '      is left exactly as it was found.'
    & $WriteLine '    * On any unexpected error the payload makes a last-resort attempt to'
    & $WriteLine '      start the Spooler, so a failure never ends with printing disabled.'
    & $WriteLine '    * Every state transition is recorded in serviceTrace[] in the report.'
    & $WriteLine ''
    & $WriteSub 'Backups'
    & $WritePara '-Backup archives the exact files it is about to delete, plus a JSON manifest listing job metadata and the SHA-256 of the archive. -BackupRequired turns a backup failure into a hard abort.'
    & $WritePara 'Honest limitation: restoring spool files into PRINTERS does NOT re-queue the job. The Spooler builds its job table in memory and only reads files it created in the current boot. A restore recovers the document payload for re-submission through the application. Anything claiming otherwise would be a lie in your runbook.'
    & $WriteLine ''
    & $WriteSub 'Confirmation'
    & $WritePara 'Interactive by default. -Force skips the prompt. On a non-interactive host (Task Scheduler, CI, redirected stdin) the tool refuses to prompt and aborts with exit code 4 rather than blocking forever - the v3 .bat hung indefinitely in exactly this situation.'
    & $WriteLine ''
    & $WriteSub 'What the tool never does'
    & $WriteLine '    * never deletes printers, ports, drivers (unless -PurgeDriverStore),'
    & $WriteLine '      or any user document outside the spool directory'
    & $WriteLine '    * never recurses into subdirectories'
    & $WriteLine '    * never touches a UNC/remote file system'
    & $WriteLine '    * never disables or reconfigures a service permanently'
    & $WriteLine '    * never writes credentials, tokens or personal data into logs/reports'
}

'Troubleshooting' {
    & $WriteHead 'TROUBLESHOOTING'
    $items = @(
        @{ S = '"The Spooler will not stop" / STOP_PENDING forever'; A = 'Almost always PrintNotify or a hung splwow64.exe holding a handle. Run with -Repair (stops dependents first) and check serviceTrace[] in the report. If it still hangs, -Force enables the process-kill fallback.' }
        @{ S = 'Files come back immediately after the cleanup'; A = 'A driver or port monitor is re-submitting them, or a client keeps resending. The report flags this as resurrectedFiles > 0. Run -Diagnostic, look at the driver and port columns, and consider -Repair -PurgeDriverStore.' }
        @{ S = '"Access is denied" on some files only'; A = 'The file is held by spoolsv.exe. The tool retries with exponential backoff (500 ms -> 1 s -> 2 s) and reports the remainder as filesLocked with the exact error.' }
        @{ S = 'Remote targets all fail with a timeout'; A = 'WinRM is not configured or 5985/5986 is blocked. Enable with Enable-PSRemoting -Force on the target and check TrustedHosts for workgroup machines. Use -PreFlightProbe to see the exact per-host reason in seconds.' }
        @{ S = 'Remote works for some hosts only'; A = 'Mixed PowerShell versions are fine (the payload is 5.1-compatible). Mixed credentials are not: pass -Credential, or ensure the caller is a local administrator on every target.' }
        @{ S = '-PrinterName matches nothing but jobs are visible'; A = 'When the Spooler is hung, WMI cannot answer and the inventory is built from .SHD headers, where the printer name is best-effort. The report says so in warnings[]. Use -Diagnostic to see the Source column (WMI vs SHD) or clean by -OlderThan instead.' }
        @{ S = 'The log file is missing'; A = 'The resolved log location is printed in the summary. If the script lives in a read-only location the tool automatically falls back to %ProgramData%, %LocalAppData% and finally %TEMP%. Pass -LogPath to force a location.' }
        @{ S = 'Update check always fails'; A = 'GitHub is unreachable (proxy/firewall) or the anonymous rate limit is hit. Set HTTPS_PROXY, supply PSQ_GITHUB_TOKEN, or disable the check with PSQ_NO_UPDATE_CHECK=true. Exit code 6 distinguishes this from "no update".' }
        @{ S = 'Update downloads but refuses to install'; A = 'The release publishes no checksums.txt. Either publish it (tools/Build.ps1 generates it) or pass -SkipChecksumVerification knowingly.' }
        @{ S = 'The task runs but does nothing'; A = 'Check the LastTaskResult with Get-SpoolerCleanupTask. 0x1 = fatal, 0x4 = it was waiting for a confirmation that can never arrive - the installer always adds -NonInteractive, so a hand-made task must add it too.' }
        @{ S = 'Garbled characters in cmd.exe output'; A = 'The tool detects the OEM code page and switches to ASCII automatically. Force it with -Ascii, or run chcp 65001 for full Unicode.' }
        @{ S = 'Spooler is fine but users still cannot print'; A = 'That is not a queue problem. Run -Diagnostic: look for WorkOffline printers, a stopped PrintNotify, a dead TCP/IP port, or a driver in an error state.' }
    )
    foreach ($it in $items) {
        & $WriteLine ''
        & $WriteLine ("  PROBLEM  {0}" -f $it.S)
        & $WriteLine '  FIX'
        foreach ($line in (Split-HelpParagraph -Text $it.A -Width ($Width - 8))) { & $WriteLine ("      $line") }
    }
    & $WriteLine ''
    & $WriteSub 'Still stuck?'
    & $WriteLine '    .\Clear-PrintSpoolerQueue.ps1 -Diagnostic -DiagnosticsBundlePath C:\Temp'
    & $WritePara 'Attach the produced ZIP to an issue. It contains the environment report, diagnostics, effective configuration with secrets redacted, the recent log buffer and file integrity hashes - everything needed to reproduce the analysis without exposing credentials.'
    & $WriteLine ("    {0}" -f $Script:ScriptIssuesUrl)
}

'Diagnostics' {
    & $WriteHead 'DIAGNOSTICS (-Diagnostic)'
    & $WritePara 'Read-only. The Spooler is never stopped and no file is ever touched, so it is safe to run on a production print server at any time.'
    & $WriteLine ''
    & $WriteSub 'What it collects'
    $what = @(
        'OS caption, version, build, last boot time, free memory',
        'Print-related services: Spooler, PrintNotify, Fax, PrintWorkflowUserSvc_* (state, start mode, PID)',
        'spoolsv.exe and splwow64.exe: CPU seconds, working set, thread count, responding flag, start time',
        'Printers: state, status, local/network/shared, offline flag, port, driver',
        'Installed printer drivers: vendor, version, INF',
        'TCP/IP printer ports: host address, port, protocol, SNMP',
        'Spool directory: file count, bytes, per-extension breakdown, oldest/newest file age, ACL owner',
        'Queued jobs: id, printer, document, owner, size, pages, status, submission time, data source (WMI or SHD)',
        'Registry: DefaultSpoolDirectory, Print event logging, driver isolation settings',
        'Event logs: Service Control Manager entries mentioning Spooler/Print, plus Microsoft-Windows-PrintService/Admin'
    )
    foreach ($w in $what) { & $WriteLine ("    * {0}" -f $w) }
    & $WriteLine ''
    & $WriteSub 'How to read it'
    & $WriteLine '    1. Spooler State=Running but jobs are not printing -> look at the printer'
    & $WriteLine '       Offline flag and the port host address.'
    & $WriteLine '    2. STOP_PENDING / STOP_FAILED -> look at splwow64.exe threads and at'
    & $WriteLine '       PrintNotify. That combination is the classic hang.'
    & $WriteLine '    3. Oldest file age >> newest -> jobs are accumulating, not draining.'
    & $WriteLine '    4. Source=SHD on every job -> the Spooler is too sick to answer WMI.'
    & $WriteLine '    5. SCM 7031/7034 events -> the service crashed; schedule an OnEvent task.'
}

'Version' {
    & $WriteHead 'VERSION'
    & $WriteLine ("    {0} v{1}" -f $Script:ScriptName, $Script:ScriptVersion)
    & $WriteLine ("    Build          : {0}" -f $Script:ScriptBuild)
    & $WriteLine ("    Product family : {0}" -f $Script:ProductFamily)
    & $WriteLine ("    Author         : {0} <{1}>" -f $Script:ScriptAuthor, $Script:ScriptAuthorEmail)
    & $WriteLine ("    Site           : {0}" -f $Script:ScriptAuthorSite)
    & $WriteLine ("    Repository     : {0}" -f $Script:ScriptRepoUrl)
    & $WriteLine ("    Wiki           : {0}" -f $Script:ScriptWikiUrl)
    & $WriteLine ("    License        : {0}" -f $Script:ScriptLicense)
    & $WriteLine ("    Host           : PowerShell {0} ({1})" -f $PSVersionTable.PSVersion, $PSVersionTable.PSEdition)
    & $WriteLine ("    Minimum        : PowerShell {0} (recommended {1}+ for -ThrottleLimit)" -f $Script:MinPSVersion, $Script:RecommendedPSVersion)
    & $WriteLine ("    OS             : {0}" -f ((Get-WindowsSku).Caption))
    & $WriteLine ("    Elevated       : {0}" -f (Test-RunningAsAdministrator))
}

'Credits' {
    & $WriteHead 'CREDITS AND LICENSE'
    & $WritePara ("Copyright (c) {0}. Licensed under {1}. See the LICENSE file in the repository for the full text." -f $Script:ScriptAuthor, $Script:ScriptLicense)
    & $WriteLine ''
    & $WriteLine '    Author    : Mikhail Deynekin'
    & $WriteLine ("    E-Mail    : {0}" -f $Script:ScriptAuthorEmail)
    & $WriteLine ("    Website   : {0}" -f $Script:ScriptAuthorSite)
    & $WriteLine ("    Source    : {0}" -f $Script:ScriptRepoUrl)
    & $WriteLine ("    Issues    : {0}" -f $Script:ScriptIssuesUrl)
    & $WriteLine ("    Docs      : {0}" -f $Script:ScriptWikiUrl)
    & $WriteLine ''
    & $WritePara 'Built for senior system administrators who would rather close the ticket permanently than restart the service every morning.'
}

        }
    }

    $text = $sb.ToString()
    if (-not [string]::IsNullOrWhiteSpace($Search)) {
        $needle = $Search.Trim()
        $kept = New-Object System.Collections.Generic.List[string]
        foreach ($block in ($text -split "(?m)^(?=={10,})")) {
            if ($block -match [regex]::Escape($needle)) { $kept.Add($block) }
        }
        if ($kept.Count -eq 0) {
            # Fall back to line-level filtering so a term inside a table still matches.
            $lines = $text -split "\r?\n" | Where-Object { $_ -match [regex]::Escape($needle) }
            if ($lines.Count -eq 0) { return "No help content matches '$needle'." }
            $text = ($lines -join $nl)
        }
        else { $text = ($kept -join $nl) }
    }
    return $text
}

function Split-HelpParagraph {
    <#
    .SYNOPSIS
        Word-wraps a paragraph to a fixed width without breaking tokens.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Text,
        [int]$Width = 96
    )
    if ([string]::IsNullOrWhiteSpace($Text)) { return @('') }
    $Width = [Math]::Max(20, $Width)
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($paragraph in ($Text -split "\r?\n")) {
        $current = ''
        foreach ($word in ($paragraph -split '\s+')) {
            if ([string]::IsNullOrEmpty($word)) { continue }
            if ($current.Length -eq 0) { $current = $word }
            elseif ($current.Length + 1 + $word.Length -le $Width) { $current += " $word" }
            else { $lines.Add($current); $current = $word }
        }
        $lines.Add($current)
    }
    return $lines.ToArray()
}

function Get-ParameterDocumentation {
    <#
    .SYNOPSIS
        Builds the parameter table from LIVE metadata of Clear-PrintSpoolerQueue.
    .DESCRIPTION
        Reading the real ParameterMetadata means the help cannot claim a switch
        that no longer exists, and cannot omit one that does. Grouping and the
        prose descriptions are supplied here; names, types, mandatory flags and
        defaults come from reflection.
    #>
    [CmdletBinding()]
    param()

    $groups = [ordered]@{
        'TARGETING'      = @('ComputerName', 'ComputerListFile', 'Credential', 'ThrottleLimit', 'PreFlightProbe', 'UseSsl')
        'SELECTION'      = @('PrinterName', 'ExcludePrinterName', 'JobId', 'DocumentName', 'OwnerName', 'OlderThan', 'NewerThan', 'LargerThanBytes', 'SmallerThanBytes', 'SpoolDirectory')
        'MODES'          = @('Diagnostic', 'Repair', 'Restore', 'CheckUpdate', 'Update', 'AutoUpdate', 'SelfTest', 'Watchdog', 'SelfTestPrint')
        'SAFETY'         = @('Force', 'Backup', 'BackupPath', 'BackupRequired', 'ArchivePath', 'RetryCount', 'RetryDelayMs', 'VerifySignature', 'ExpectedSignerThumbprint', 'NonInteractive', 'WatchdogMinutes')
        'REPAIR OPTIONS' = @('PurgeFormsCache', 'PurgeDriverStore', 'ResetSpoolDirectory', 'StopDependentServices', 'RestartDependentServices')
        'OUTPUT'         = @('Quiet', 'NoColor', 'Ascii', 'NoProgress', 'MachineReadableJson', 'ExportFormat', 'ExportPath', 'PassThru', 'DiagnosticsBundlePath')
        'LOGGING'        = @('LogPath', 'LogLevel', 'JsonLog', 'EventLog', 'LogRetentionDays', 'NoLogRotation')
        'CONFIGURATION'  = @('ConfigPath', 'SaveConfig', 'NoConfigFile', 'ServiceTimeoutSeconds', 'SettleSeconds')
        'UPDATE'         = @('UpdateChannel', 'IncludePrerelease', 'SkipChecksumVerification', 'GitHubToken', 'NoUpdateCheck', 'ApplyUpdateNow')
        'ENVIRONMENT'    = @('SelfElevate', 'NoElevate', 'NoProfile')
    }

    $descriptions = @{
        ComputerName              = @('One or more target computers. Defaults to the local machine. Accepts pipeline input (by value and by property name), so Get-ADComputer / Import-Csv / Get-Content all feed it directly.', 'Names are de-duplicated and empty entries are skipped. localhost, ., 127.0.0.1, ::1 and the local host name all mean "this machine".')
        ComputerListFile          = @('Path to a text file with one computer name per line. Blank lines and # comments are ignored. Merged with -ComputerName.', $null)
        Credential                = @('PSCredential used for remote connections. Never written to logs, reports or the diagnostics bundle.', $null)
        ThrottleLimit             = @('Number of remote targets processed concurrently. 1 = sequential (the safe default).', 'Requires PowerShell 7 (ForEach-Object -Parallel). On Windows PowerShell 5.1 an equivalent runspace pool is used automatically.')
        PreFlightProbe            = @('Probe WinRM on every remote target before doing real work, so offline hosts fail in ~1 s instead of ~40 s.', $null)
        UseSsl                    = @('Use HTTPS (WinRM 5986) instead of HTTP (5985) for remote sessions.', $null)
        PrinterName               = @('Only act on jobs for these printers. Wildcards allowed ("HP*", "*Kyocera*"). Case-insensitive.', 'When a job has no readable printer metadata (typical when the Spooler is hung and the inventory comes from .SHD headers) it is NOT selected, and the report says so. Use no filter when the whole queue must go.')
        ExcludePrinterName        = @('Never act on jobs for these printers. Applied after -PrinterName.', $null)
        JobId                     = @('Act only on these job ids, as shown by -Diagnostic or Get-PrintJob.', $null)
        DocumentName              = @('Filter on the document name with wildcards, e.g. "*invoice*.pdf".', $null)
        OwnerName                 = @('Filter on the submitting user, e.g. "CORP\jdoe".', $null)
        OlderThan                 = @('Only act on jobs submitted longer ago than this. Accepts 30s, 15m, 2h, 7d or a [TimeSpan].', 'The single most useful filter in production: it clears the stuck jobs and leaves the one that is legitimately printing right now.')
        NewerThan                 = @('Only act on jobs submitted more recently than this.', $null)
        LargerThanBytes           = @('Only act on jobs bigger than this. Accepts a byte count or "10MB"/"1.5GB".', $null)
        SmallerThanBytes          = @('Only act on jobs smaller than this.', $null)
        RestoreDestination        = @('With -Restore: directory to extract into. Default: a timestamped "restored_<computer>_<stamp>" folder next to the archive.', $null)
        EventLogHours             = @('With -Diagnostic: how many hours of event-log history to collect. Default 72.', $null)
        Help                      = @('Render the complete built-in reference and exit.', 'A switch: bare -Help works. For a single section use -HelpSection <name>, or -ListHelpSections to see the names.')
        HelpSection               = @('Render just this section of the built-in reference and exit.', 'Sections: All, Usage, QuickStart, Modes, Parameters, Examples, ExitCodes, Configuration, Environment, Logging, Reports, Remote, Scheduling, Updates, Safety, Troubleshooting, Diagnostics, Version, Credits.')
        ListHelpSections          = @('List the available help section names (for -HelpSection) and exit.', $null)
        ShowVersion               = @('Print version, build, host, OS, elevation and integrity facts, then exit.', 'With -PassThru the same facts come back as an object, which is what an asset inventory should collect.')
        SpoolDirectory            = @('Override the spool directory instead of resolving it from the registry.', 'Still subject to every safety guard. Intended for hosts with an unusual layout, not for pointing the tool at an arbitrary folder.')
        Diagnostic                = @('Read-only mode: collect a full health and queue inventory and change nothing.', 'Safe on production at any time. Exit code 11 means "diagnostics completed".')
        Repair                    = @('Deep clean: cleanup plus root-cause repair steps (dependent-service cycling, Forms cache, optional driver-store purge, optional spool-directory reset).', $null)
        Restore                   = @('Path to a backup ZIP created by -Backup. Restores it forensically and exits.', 'Copying spool files back into PRINTERS does not re-queue jobs; restore to a folder and re-submit the documents.')
        CheckUpdate               = @('Ask GitHub whether a newer version exists, report it and exit. Changes nothing.', 'Exit code 7 = update available, 6 = could not reach GitHub, 0 = already current.')
        Update                    = @('Download, verify and install the newest version, then exit.', $null)
        AutoUpdate                = @('Perform a quiet, cooldown-limited update check before the cleanup and stage a new version if one exists.', 'Designed for scheduled tasks: it never blocks, never fails the run, and checks at most once per cooldown window.')
        SelfTest                  = @('Run the built-in consistency checks (payload generation, guards, schema, version parsing) and exit.', $null)
        Watchdog                  = @('Arm the watchdog: if the run exceeds -WatchdogMinutes, restart the Spooler and terminate.', $null)
        SelfTestPrint             = @('After cleanup, submit and immediately cancel a test job to prove the queue really works again.', 'Turns "the service says Running" into "printing actually works".')
        Force                     = @('Skip the interactive confirmation prompt. Required for unattended execution.', $null)
        Backup                    = @('Archive every file that is about to be deleted, with a JSON manifest and SHA-256, before deleting anything.', $null)
        BackupPath                = @('Backup destination directory. Default: %ProgramData%\PrintSpoolerCleanup\Backups.', $null)
        BackupRequired            = @('Abort the cleanup if the backup could not be created.', 'Recommended wherever a print queue has legal or financial significance.')
        ArchivePath               = @('Move (not copy) the created backup archives here after the run, e.g. a log share.', $null)
        RetryCount                = @('Delete attempts per file before it is reported as locked. Backoff is exponential.', $null)
        RetryDelayMs              = @('Initial delay between delete retries in milliseconds (500 -> 1000 -> 2000 by default).', $null)
        VerifySignature           = @('Verify the Authenticode signature of the running script before doing anything else.', 'Exit code 14 on a broken, missing or unexpected signature.')
        ExpectedSignerThumbprint  = @('Require the signature to come from this certificate thumbprint.', $null)
        NonInteractive            = @('Never prompt. If confirmation would be required, abort with exit code 4 instead of waiting.', $null)
        WatchdogMinutes           = @('Watchdog limit in minutes. 0 disables it.', $null)
        PurgeFormsCache           = @('With -Repair: clear the spooler Forms cache, which a corrupt custom form can wedge.', $null)
        PurgeDriverStore          = @('With -Repair: remove printer driver packages that no printer references.', 'Destructive and therefore requires -Force as an explicit acknowledgement. In use only by an active printer is never removed.')
        ResetSpoolDirectory       = @('With -Repair: remove the DefaultSpoolDirectory registry override so the OS default is used again.', $null)
        StopDependentServices     = @('Stop Fax / PrintNotify / PrintWorkflowUserSvc_* before the Spooler. Default: on.', 'This is what makes the stop succeed on hosts where v3 hung.')
        RestartDependentServices  = @('Restart the dependent services that were running before. Default: on.', 'Services that were stopped or disabled are left exactly as they were.')
        Quiet                     = @('Suppress all decorative console output. Logging and reports are unaffected.', $null)
        NoColor                   = @('Plain monochrome output (also triggered by NO_COLOR or a redirected stream).', $null)
        Ascii                     = @('Use ASCII glyphs instead of Unicode. Auto-enabled when the console code page cannot render them.', $null)
        NoProgress                = @('Suppress Write-Progress bars (useful in CI and in constrained hosts).', $null)
        MachineReadableJson       = @('Write exactly one JSON document to stdout and nothing else.', 'This is the contract for RMM/monitoring ingestion.')
        ExportFormat              = @('Report formats to write: Json, Csv, Html, Text.', $null)
        ExportPath                = @('Report destination: a directory (auto-named files) or a full file path.', $null)
        PassThru                  = @('Return the run summary object to the pipeline in addition to rendering it.', $null)
        DiagnosticsBundlePath     = @('Also write a ZIP bundle suitable for attaching to a support issue.', $null)
        LogPath                   = @('Explicit log file. The parent directory is created if needed.', 'If the location is not writable the tool falls back automatically and says so.')
        LogLevel                  = @('Minimum level recorded: TRACE, DEBUG, INFO, WARN, ERROR, CRITICAL.', $null)
        JsonLog                   = @('Write the log as JSON Lines instead of pipe-delimited text.', $null)
        EventLog                  = @('Also write to the Windows Event Log (source: PrintSpoolerCleanup).', 'Creating the source needs elevation; failure disables event logging for the run only.')
        LogRetentionDays          = @('Delete rotated logs older than this many days. 0 disables pruning.', $null)
        NoLogRotation             = @('Disable size-based log rotation for this run.', $null)
        ConfigPath                = @('Load configuration from this .psd1 file.', 'Parsed with the AST and SafeGetValue(); executable content is rejected.')
        SaveConfig                = @('Write the effective configuration to -ConfigPath and exit.', $null)
        NoConfigFile              = @('Ignore any auto-discovered configuration file.', $null)
        ServiceTimeoutSeconds     = @('Seconds to wait for a service to reach the desired state.', $null)
        SettleSeconds             = @('Pause after the Spooler stops, so handles are released before deleting.', $null)
        UpdateChannel             = @('latest, stable or prerelease.', $null)
        IncludePrerelease         = @('Treat prereleases as valid update targets.', $null)
        SkipChecksumVerification  = @('Install an update even when the release publishes no checksums.txt.', 'Deliberately explicit: the default is to refuse unverified bytes.')
        GitHubToken               = @('Token for the GitHub API (raises the anonymous 60 req/h limit to 5000).', 'Prefer the PSQ_GITHUB_TOKEN environment variable so the token never appears in a command line or process list.')
        NoUpdateCheck             = @('Disable every automatic update check for this run.', $null)
        ApplyUpdateNow            = @('With -Update/-AutoUpdate: move the staged files into place immediately instead of on next start.', $null)
        SelfElevate               = @('Re-launch elevated through UAC when not running as administrator.', $null)
        NoElevate                 = @('Never attempt elevation; fail with exit code 5 instead.', $null)
        NoProfile                 = @('Ignored for compatibility with launcher scripts that always pass it.', $null)
    }

    $docs = New-Object System.Collections.Generic.List[object]
    $cmd = Get-Command -Name 'Clear-PrintSpoolerQueue' -ErrorAction SilentlyContinue
    $params = @()
    if ($cmd) {
        $params = @($cmd.Parameters.GetEnumerator() | Where-Object {
            $_.Key -notin @('Verbose', 'Debug', 'ErrorAction', 'WarningAction', 'InformationAction',
                            'ErrorVariable', 'WarningVariable', 'InformationVariable', 'OutVariable',
                            'OutBuffer', 'PipelineVariable', 'ProgressAction', 'Confirm', 'WhatIf')
        })
    }

    $emitted = New-Object System.Collections.Generic.List[string]
    foreach ($group in $groups.Keys) {
        foreach ($name in $groups[$group]) {
            $meta = $params | Where-Object { $_.Key -eq $name } | Select-Object -First 1
            $type = 'switch'
            $default = $null
            if ($meta) {
                $t = $meta.Value.ParameterType
                $type = if ($t -eq [switch]) { 'switch' }
                        elseif ($t -eq [string[]]) { 'string[]' }
                        elseif ($t -eq [string]) { 'string' }
                        elseif ($t -eq [int]) { 'int' }
                        elseif ($t -eq [long]) { 'long' }
                        else { $t.Name }
                $attr = @($meta.Value.Attributes | Where-Object { $_ -is [System.Management.Automation.ParameterAttribute] })
                $mandatory = @($attr | Where-Object { $_.Mandatory }).Count -gt 0
                if ($mandatory) { $type = "$type (mandatory)" }
            }
            $desc = if ($descriptions.Contains($name)) { $descriptions[$name][0] } else { '' }
            $note = if ($descriptions.Contains($name)) { $descriptions[$name][1] } else { $null }
            $docs.Add([pscustomobject]@{ Group = $group; Name = $name; Type = $type; Default = $default; Description = $desc; Note = $note })
            $emitted.Add($name) | Out-Null
        }
    }
    # Any parameter that exists in code but is missing from the group map is
    # still documented - an undocumented switch is a defect, not a feature.
    foreach ($meta in $params) {
        if ($emitted -contains $meta.Key) { continue }
        $t = $meta.Value.ParameterType
        $type = if ($t -eq [switch]) { 'switch' } else { $t.Name }
        $desc = if ($descriptions.Contains($meta.Key)) { $descriptions[$meta.Key][0] } else { 'No description available (report this as a documentation defect).' }
        $docs.Add([pscustomobject]@{ Group = 'OTHER'; Name = $meta.Key; Type = $type; Default = $null; Description = $desc; Note = $null })
    }
    return $docs.ToArray()
}

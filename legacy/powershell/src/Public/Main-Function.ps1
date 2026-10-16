# ============================================================================
#  Clear-PrintSpoolerQueue  -  main entry point
# ----------------------------------------------------------------------------
#  Orchestration only. Every substantive operation lives in a private helper so it
#  can be unit-tested without touching a real Spooler:
#
#    payload      06-RemotePayload.ps1   (runs ON the target machine)
#    dispatch     07-Orchestration.ps1   (local / remote / parallel)
#    config       05-Configuration.ps1
#    logging      02-Logging.ps1
#    reporting    10-Report.ps1
#    updates      09-Update.ps1
#    watchdog     14-Execution.ps1
#
#  Execution order matters and is deliberate:
#    1. state + palette + log target   (so every later step can be logged)
#    2. pending-update swap            (before running any more of our own code)
#    3. read-only modes (help / version / update check / restore / self-test)
#    4. platform, elevation, signature (fail before touching a print server)
#    5. configuration + parameter validation
#    6. confirmation gate
#    7. watchdog arming
#    8. dispatch
#    9. aggregate, report, export, exit code
# ============================================================================

function Clear-PrintSpoolerQueue {
    <#
    .SYNOPSIS
        Clears stuck Windows print queues safely, selectively and verifiably -
        locally or across a fleet - with backup, diagnostics, repair, structured
        logging, machine-readable reports and built-in self-updating.

    .DESCRIPTION
        The Print Spooler holds queued jobs as *.SHD (shadow/metadata) and *.SPL
        (spool data) files in %SystemRoot%\System32\spool\PRINTERS. When a job
        cannot be rendered, when a driver faults, or when a port monitor blocks,
        the queue wedges: the service reports Running, jobs will not cancel, and
        everything queued behind them stalls.

        The canonical fix - stop the Spooler, delete the spool files, start the
        Spooler - is correct but naive. In practice it fails because:
          * dependent services (Fax, PrintNotify, PrintWorkflowUserSvc_*) hold
            handles, so the stop hangs in STOP_PENDING
          * the files are locked by spoolsv.exe, and one locked file aborts the run
          * deleting the WHOLE queue destroys jobs that were printing fine
          * there is no undo once the files are gone
          * nothing verifies that printing actually works afterwards

        This tool addresses each of those. It stops dependents first, retries
        locked files with exponential backoff, can select jobs by printer, job id,
        owner, document name, age or size, archives everything it is about to
        delete, restarts only the services that were running, verifies the queue
        did not immediately refill, and reports a machine-readable result.

        One implementation runs on every target: the same payload executes locally
        and through WinRM, so local and remote behaviour cannot drift.

    .PARAMETER ComputerName
        Target computers. Defaults to the local machine. Accepts pipeline input by
        value and by property name, so Get-ADComputer, Import-Csv and Get-Content
        all feed it directly.

    .PARAMETER Force
        Skip the interactive confirmation. Required for unattended execution.

    .PARAMETER Backup
        Archive every file that is about to be deleted (ZIP + JSON manifest +
        SHA-256) before deleting anything.

    .PARAMETER OlderThan
        Only act on jobs submitted longer ago than this. Accepts 30s, 15m, 2h, 7d
        or a [TimeSpan]. The most useful production filter: it clears the wedged
        jobs and leaves the one that is legitimately printing right now.

    .PARAMETER Diagnostic
        Read-only deep inspection. Never stops the service, never deletes a file.

    .PARAMETER Repair
        Cleanup plus root-cause repair: dependent-service cycling, Forms cache
        purge, optional driver-store purge, optional spool-directory reset.

    .PARAMETER Help
        Render the complete built-in reference and exit. Combine with
        -HelpSection to render a single section instead.

    .PARAMETER HelpSection
        Render just this section of the built-in reference and exit. Section
        names: All, Usage, QuickStart, Modes, Parameters, Examples, ExitCodes,
        Configuration, Environment, Logging, Reports, Remote, Scheduling,
        Updates, Safety, Troubleshooting, Diagnostics, Version, Credits.
        -ListHelpSections prints the same names.

    .EXAMPLE
        .\Clear-PrintSpoolerQueue.ps1 -WhatIf
        Shows exactly what would happen and changes nothing.

    .EXAMPLE
        .\Clear-PrintSpoolerQueue.ps1 -Force -Backup
        Clears the local queue with an archive of everything removed.

    .EXAMPLE
        .\Clear-PrintSpoolerQueue.ps1 -Diagnostic -ExportFormat Html -ExportPath C:\Temp
        Produces a self-contained HTML health report for a ticket.

    .EXAMPLE
        .\Clear-PrintSpoolerQueue.ps1 -ComputerName (Get-Content .\hosts.txt) -ThrottleLimit 16 -Force -Backup
        Fleet-wide maintenance in parallel (PowerShell 7).

    .EXAMPLE
        .\Clear-PrintSpoolerQueue.ps1 -OlderThan 2h -PrinterName "HP*" -Force
        Only HP jobs older than two hours.

    .EXAMPLE
        .\Clear-PrintSpoolerQueue.ps1 -Force -NonInteractive -Quiet -MachineReadableJson -Backup
        Unattended run whose stdout is a single JSON document for an RMM.

    .EXAMPLE
        .\Clear-PrintSpoolerQueue.ps1 -Help
        The complete built-in reference.

    .EXAMPLE
        .\Clear-PrintSpoolerQueue.ps1 -HelpSection Troubleshooting
        One section of the reference: symptom -> cause -> fix.

    .OUTPUTS
        None by default. With -PassThru, a run summary object whose Results
        collection holds one record per computer.

    .NOTES
        Author  : Mikhail Deynekin <Mikhail@Deynekin.com>
        Site    : https://Deynekin.com
        Repo    : https://github.com/paulmann/Print-Spooler-Queue-Cleanup
        License : GPL-3.0-or-later

        Requires Windows PowerShell 5.1 or PowerShell 7+. Administrator rights for
        the local machine; WinRM plus authorization for remote targets.
        -ThrottleLimit > 1 needs PowerShell 7 (a runspace pool is used on 5.1).
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    [OutputType([pscustomobject])]
    param(
        # ---------------------------------------------------------------- targeting
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
        [Alias('Computer', 'CN', 'HostName')]
        [AllowNull()][AllowEmptyCollection()][string[]]$ComputerName,

        [Parameter()]
        [Alias('HostsFile', 'TargetsFile')]
        [AllowNull()][AllowEmptyString()][string]$ComputerListFile,

        [Parameter()]
        [Alias('RunAs')]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter()]
        [ValidateRange(1, 64)]
        [int]$ThrottleLimit = 1,

        [Parameter()]
        [bool]$PreFlightProbe = $true,

        [Parameter()]
        [switch]$UseSsl,

        # ---------------------------------------------------------------- selection
        [Parameter()]
        [Alias('Printer', 'Queue')]
        [AllowNull()][AllowEmptyCollection()][string[]]$PrinterName,

        [Parameter()]
        [AllowNull()][AllowEmptyCollection()][string[]]$ExcludePrinterName,

        [Parameter()]
        [Alias('Id')]
        [AllowNull()][AllowEmptyCollection()][long[]]$JobId,

        [Parameter()]
        [AllowNull()][AllowEmptyString()][string]$DocumentName,

        [Parameter()]
        [Alias('Owner', 'SubmittedBy')]
        [AllowNull()][AllowEmptyString()][string]$OwnerName,

        [Parameter()]
        [Alias('MinAge')]
        [AllowNull()]$OlderThan,

        [Parameter()]
        [Alias('MaxAge')]
        [AllowNull()]$NewerThan,

        [Parameter()]
        [Alias('MinSize')]
        [AllowNull()]$LargerThanBytes,

        [Parameter()]
        [Alias('MaxSize')]
        [AllowNull()]$SmallerThanBytes,

        [Parameter()]
        [AllowNull()][AllowEmptyString()][string]$SpoolDirectory,

        # ---------------------------------------------------------------- modes
        [Parameter()]
        [switch]$Diagnostic,

        [Parameter()]
        [switch]$Repair,

        [Parameter()]
        [AllowNull()][AllowEmptyString()][string]$Restore,

        [Parameter()]
        [AllowNull()][AllowEmptyString()][string]$RestoreDestination,

        [Parameter()]
        [switch]$CheckUpdate,

        [Parameter()]
        [switch]$Update,

        [Parameter()]
        [switch]$AutoUpdate,

        [Parameter()]
        [switch]$SelfTest,

        [Parameter()]
        [switch]$Watchdog,

        [Parameter()]
        [switch]$SelfTestPrint,

        # ---------------------------------------------------------------- safety
        [Parameter()]
        [switch]$Force,

        [Parameter()]
        [switch]$Backup,

        [Parameter()]
        [AllowNull()][AllowEmptyString()][string]$BackupPath,

        [Parameter()]
        [switch]$BackupRequired,

        [Parameter()]
        [AllowNull()][AllowEmptyString()][string]$ArchivePath,

        [Parameter()]
        [ValidateRange(0, 20)]
        [int]$RetryCount = 3,

        [Parameter()]
        [ValidateRange(0, 60000)]
        [int]$RetryDelayMs = 500,

        [Parameter()]
        [switch]$VerifySignature,

        [Parameter()]
        [AllowNull()][AllowEmptyString()][string]$ExpectedSignerThumbprint,

        [Parameter()]
        [Alias('NoPrompt')]
        [switch]$NonInteractive,

        [Parameter()]
        [ValidateRange(0, 1440)]
        [int]$WatchdogMinutes = 30,

        # ---------------------------------------------------------------- repair options
        [Parameter()]
        [switch]$PurgeFormsCache,

        [Parameter()]
        [switch]$PurgeDriverStore,

        [Parameter()]
        [switch]$ResetSpoolDirectory,

        [Parameter()]
        [bool]$StopDependentServices = $true,

        [Parameter()]
        [bool]$RestartDependentServices = $true,

        # ---------------------------------------------------------------- output
        [Parameter()]
        [Alias('Silent')]
        [switch]$Quiet,

        [Parameter()]
        [switch]$NoColor,

        [Parameter()]
        [Alias('NoUnicode')]
        [switch]$Ascii,

        [Parameter()]
        [switch]$NoProgress,

        [Parameter()]
        [Alias('Json')]
        [switch]$MachineReadableJson,

        [Parameter()]
        [ValidateSet('Json', 'Csv', 'Html', 'Text')]
        [Alias('ReportFormat')]
        [AllowNull()][AllowEmptyCollection()][string[]]$ExportFormat,

        [Parameter()]
        [Alias('ReportPath')]
        [AllowNull()][AllowEmptyString()][string]$ExportPath,

        [Parameter()]
        [switch]$PassThru,

        [Parameter()]
        [AllowNull()][AllowEmptyString()][string]$DiagnosticsBundlePath,

        # ---------------------------------------------------------------- logging
        [Parameter()]
        [Alias('Log')]
        [AllowNull()][AllowEmptyString()][string]$LogPath,

        [Parameter()]
        [ValidateSet('TRACE', 'DEBUG', 'INFO', 'WARN', 'ERROR', 'CRITICAL')]
        [AllowNull()][AllowEmptyString()][string]$LogLevel,

        [Parameter()]
        [switch]$JsonLog,

        [Parameter()]
        [switch]$EventLog,

        [Parameter()]
        [ValidateRange(0, 3650)]
        [int]$LogRetentionDays = 30,

        [Parameter()]
        [switch]$NoLogRotation,

        # ---------------------------------------------------------------- configuration
        [Parameter()]
        [Alias('Config')]
        [AllowNull()][AllowEmptyString()][string]$ConfigPath,

        [Parameter()]
        [switch]$SaveConfig,

        [Parameter()]
        [switch]$NoConfigFile,

        [Parameter()]
        [Alias('TimeoutSeconds')]
        [ValidateRange(5, 3600)]
        [int]$ServiceTimeoutSeconds = 60,

        [Parameter()]
        [ValidateRange(0, 300)]
        [int]$SettleSeconds = 2,

        [Parameter()]
        [ValidateRange(0, 2160)]
        [int]$EventLogHours = 72,

        # ---------------------------------------------------------------- updates
        [Parameter()]
        [ValidateSet('latest', 'stable', 'prerelease')]
        [string]$UpdateChannel = 'latest',

        [Parameter()]
        [switch]$IncludePrerelease,

        [Parameter()]
        [switch]$SkipChecksumVerification,

        [Parameter()]
        [AllowNull()][AllowEmptyString()][string]$GitHubToken,

        [Parameter()]
        [switch]$NoUpdateCheck,

        [Parameter()]
        [switch]$ApplyUpdateNow,

        # ---------------------------------------------------------------- environment
        [Parameter()]
        [Alias('Elevate')]
        [switch]$SelfElevate,

        [Parameter()]
        [switch]$NoElevate,

        [Parameter()]
        [switch]$NoProfile,

        # ---------------------------------------------------------------- help
        # -Help is a SWITCH, not a string (defect 46). A named [string] parameter
        # can never be invoked bare: `-Help` alone is the binding error "Missing
        # an argument for parameter 'Help'" - yet bare `-Help` is the form the
        # README, the launcher hint and this file's own examples advertise. The
        # section selector therefore lives in -HelpSection; bare -Help renders
        # the complete reference.
        [Parameter()]
        [switch]$Help,

        [Parameter()]
        [ValidateSet('All', 'Usage', 'QuickStart', 'Modes', 'Parameters', 'Examples',
            'ExitCodes', 'Configuration', 'Environment', 'Logging', 'Reports',
            'Remote', 'Scheduling', 'Updates', 'Safety', 'Troubleshooting',
            'Diagnostics', 'Version', 'Credits')]
        [AllowNull()][AllowEmptyString()][string]$HelpSection,

        [Parameter()]
        [switch]$ListHelpSections,

        [Parameter()]
        [switch]$ShowVersion
    )

    begin {
        $Script:RuntimeState = $null
        $null = Initialize-RuntimeState
        $state = $Script:RuntimeState

        # ---- the console palette must exist before the first message is rendered
        $capability = Get-ConsoleCapability
        $null = Initialize-ConsolePalette -Ascii:$Ascii -NoColor:($NoColor -or $MachineReadableJson) -Capability $capability
        $state.Quiet = [bool]($Quiet -or $MachineReadableJson)
        # PSQ_NONINTERACTIVE is documented in -HelpSection Environment and in the wiki, so
        # it must actually work. It was dead: an operator who set it for a
        # scheduled task or a CI job got an interactive confirmation prompt anyway,
        # and with no console attached the run HUNG until the task timeout instead
        # of failing fast. Precedence: -NonInteractive > PSQ_NONINTERACTIVE >
        # automatic host detection.
        $state.NonInteractive = [bool]($NonInteractive -or
            (ConvertFrom-EnvironmentBoolean $env:PSQ_NONINTERACTIVE) -or
            (Test-NonInteractiveHost))
        $state.NoProgress = [bool]($NoProgress -or $MachineReadableJson -or $state.NonInteractive)
        $state.MachineReadable = [bool]$MachineReadableJson
        $state.ShowTimestamp = $true

        $state.Mode = 'Cleanup'
        if ($Diagnostic) { $state.Mode = 'Diagnostic' }
        if ($Repair) { $state.Mode = 'Repair' }
        if ($Restore) { $state.Mode = 'Restore' }
        if ($SelfTest) { $state.Mode = 'SelfTest' }
        if ($Help -or $HelpSection -or $ListHelpSections) { $state.Mode = 'Help' }
        if ($ShowVersion) { $state.Mode = 'Version' }
        if ($SaveConfig) { $state.Mode = 'SaveConfig' }
        if ($CheckUpdate -and -not $Update) { $state.Mode = 'CheckUpdate' }
        if ($Update) { $state.Mode = 'Update' }

        # ---- configuration is loaded HERE, not in step 4, because the log
        # target depends on it. Resolving the log from the -LogPath parameter
        # alone (as an earlier revision did) silently ignored BOTH the documented
        # PSQ_LOG_PATH environment variable and Logging.Path from the config file,
        # so a site-wide log location never took effect and every run fell back to
        # auto-discovery. Precedence: -LogPath > PSQ_LOG_PATH > Logging.Path > auto.
        # PSQ_CONFIG_PATH is documented, so it must resolve. Precedence mirrors the
        # log path: an explicit -ConfigPath wins, then the environment variable, then
        # auto-discovery. Without this, a site that pointed PSQ_CONFIG_PATH at a
        # machine-wide policy file silently ran on built-in defaults.
        $effectiveConfigPath = $ConfigPath
        if ([string]::IsNullOrWhiteSpace($effectiveConfigPath)) { $effectiveConfigPath = $env:PSQ_CONFIG_PATH }

        $earlyConfig = Get-EffectiveConfiguration -ConfigPath $effectiveConfigPath -ScriptDirectory $state.ScriptDirectory -NoConfigFile:$NoConfigFile
        $state.Configuration = $earlyConfig.Config
        $state.ConfigSource = $earlyConfig.Source
        $state.ConfigPath = $earlyConfig.Path
        $state.ConfigProblems = $earlyConfig.Problems

        $effectiveLogPath = $LogPath
        if ([string]::IsNullOrWhiteSpace($effectiveLogPath)) { $effectiveLogPath = $env:PSQ_LOG_PATH }
        if ([string]::IsNullOrWhiteSpace($effectiveLogPath)) { $effectiveLogPath = [string]$earlyConfig.Config['Logging']['Path'] }

        $stampLog = ($state.Mode -notin @('Help', 'Version', 'CheckUpdate', 'SelfTest'))
        $logTarget = Resolve-LogTarget -ExplicitPath $effectiveLogPath -SessionStamped:$stampLog
        if ($logTarget.Writable) {
            $state.LogPath = $logTarget.Path
            $state.LogSource = $logTarget.Source
        }
        else {
            $state.LogPath = $null
            if (-not $state.Quiet -and $logTarget.Error) {
                Write-Warning ("Logging disabled: {0}" -f $logTarget.Error)
            }
        }

        # ---- Ctrl+C: try to leave the Spooler running rather than stopped
        try {
            $null = Register-EngineEvent -SourceIdentifier 'ConsoleCancelEventArgs' -Action {
                try { Start-Service -Name 'Spooler' -ErrorAction SilentlyContinue }
                catch { /* the service may already be running */ }
            }
        }
        catch {
            # Hosts without a console cancel event simply do not get the handler.
        }
    }

    process {
        # Pipeline targets accumulate here; the work happens in end{} so that a
        # piped list is processed as ONE run with ONE summary and ONE exit code.
        foreach ($c in @($ComputerName)) {
            if ([string]::IsNullOrWhiteSpace($c)) { continue }
            $t = $c.Trim()
            if ($state.TargetList -notcontains $t) { $state.TargetList.Add($t) }
        }
    }

    end {
        $state = $Script:RuntimeState
        $state.RunStartedUtc = [DateTime]::UtcNow
        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        try {
            # ================================================================
            #  STEP 1 - log level, header, pending-update swap
            # ================================================================
            $effectiveLevel = if ($LogLevel) { $LogLevel } else { $Script:DefaultLogLevel }
            $null = Set-RuntimeLogLevel -Level $effectiveLevel
            $state.JsonLog = [bool]$JsonLog
            $state.EventLogEnabled = [bool]$EventLog

            $modeNotes = New-Object System.Collections.Generic.List[string]
            if ($WhatIfPreference) { $modeNotes.Add('WhatIf / dry-run - no changes will be made') }
            if ($Diagnostic) { $modeNotes.Add('Diagnostic - read-only') }
            if ($Repair) { $modeNotes.Add('Repair - deep clean') }
            if ($MachineReadableJson) { $modeNotes.Add('Machine-readable JSON on stdout') }
            Write-ConsoleHeader -ModeNotes $modeNotes.ToArray()

            Write-LogHeader -Context @{
                Mode       = $state.Mode
                Invocation = $state.InvocationId
                PSVersion  = $PSVersionTable.PSVersion.ToString()
                Elevated   = (Test-RunningAsAdministrator)
                Log        = $state.LogPath
            }
            Write-LogInfo ("v{0} started (mode={1}, invocation={2}, PowerShell {3} {4})." -f `
                $Script:ScriptVersion, $state.Mode, $state.InvocationId, $PSVersionTable.PSEdition, $PSVersionTable.PSVersion) -Component 'Session'

            # Apply a previously staged update BEFORE running any more of our own
            # code. A process cannot safely replace the file it is executing, so the
            # swap always happens at start-up, never mid-run.
            try {
                $pending = Invoke-PendingUpdateInternal -InstallDirectory $state.InstallDirectory
                if ($pending.Applied) {
                    Write-LogSuccess ("Applied a staged update to v{0} ({1}). Restart the tool to run the new version." -f `
                        $pending.ToVersion, (@($pending.Files) -join ', ')) -Component 'Update'
                }
                elseif ($pending.Message -and ($pending.Message -ne 'No pending update.')) {
                    Write-LogDebug $pending.Message -Component 'Update'
                }
            }
            catch {
                Write-LogDebug "Pending-update check skipped: $($_.Exception.Message)" -Component 'Update'
            }

            # ================================================================
            #  STEP 2 - read-only and self-contained modes
            # ================================================================
            if ($ListHelpSections) {
                foreach ($s in (Get-HelpSectionNames)) { Write-Host $s }
                $null = Set-ScriptExitCode $Script:ExitCode.Success
                if ($PassThru) { return (Get-HelpSectionNames) }
                return
            }

            if ($Help -or $HelpSection) {
                $sectionToRender = if ([string]::IsNullOrWhiteSpace($HelpSection)) { 'All' } else { $HelpSection }
                $text = Get-HelpText -Section $sectionToRender -IncludeParameterMetadata
                if ($state.MachineReadable) { Write-Output ($text | ConvertTo-Json -Compress) }
                else { Write-Host $text }
                $null = Set-ScriptExitCode $Script:ExitCode.Success
                return
            }

            if ($ShowVersion) {
                $v = Get-SpoolerCleanupVersion -Detailed
                if ($state.MachineReadable) { ($v | ConvertTo-Json -Depth 6) | Write-Output }
                else {
                    Write-ConsoleKeyValue -Label 'Tool' -Value ("{0} v{1}" -f $Script:ScriptName, $Script:ScriptVersion)
                    Write-ConsoleKeyValue -Label 'Build' -Value $Script:ScriptBuild
                    Write-ConsoleKeyValue -Label 'Repository' -Value $Script:ScriptRepoUrl
                    Write-ConsoleKeyValue -Label 'License' -Value $Script:ScriptLicense
                    Write-ConsoleKeyValue -Label 'Host' -Value ("PowerShell {0} ({1})" -f $PSVersionTable.PSVersion, $PSVersionTable.PSEdition)
                    Write-ConsoleKeyValue -Label 'OS' -Value ((Get-WindowsSku).Caption)
                    Write-ConsoleKeyValue -Label 'Elevated' -Value (Test-RunningAsAdministrator)
                    Write-ConsoleKeyValue -Label 'Install dir' -Value $state.InstallDirectory
                }
                $null = Set-ScriptExitCode $Script:ExitCode.Success
                if ($PassThru) { return $v }
                return
            }

            if ($SelfTest) {
                # The result is captured with $null = on the non-returning path.
                # An emitted value in statement position binds to the NEXT
                # command's parameters, and -PassThru is a switch - which is how a
                # self-test object ended up being converted to SwitchParameter.
                $selfTestResult = Invoke-SelfTestInternal
                if ($state.MachineReadable) { ($selfTestResult | ConvertTo-Json -Depth 6) | Write-Output }
                else {
                    Write-ConsoleSection -Title 'SELF TEST' -Character '='
                    @($selfTestResult.Checks) | Write-ConsoleTable -Property @('Name', 'Result', 'Detail')
                    $color = if ($selfTestResult.Overall -eq 'Passed') { 'Green' } else { 'Red' }
                    Write-ConsoleKeyValue -Label 'Overall' -Value ("{0} ({1}/{2})" -f $selfTestResult.Overall, ($selfTestResult.TotalChecks - $selfTestResult.FailedChecks), $selfTestResult.TotalChecks) -ValueColor $color
                }
                $code = if ($selfTestResult.Overall -eq 'Passed') { $Script:ExitCode.Success } else { $Script:ExitCode.Failed }
                $null = Set-ScriptExitCode $code
                if ($PassThru) { return $selfTestResult }
                return
            }

            if ($SaveConfig) {
                if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
                    Write-LogError '-SaveConfig requires -ConfigPath <file>.' -Component 'Config'
                    $null = Set-ScriptExitCode $Script:ExitCode.BadArguments
                    return
                }
                $effective = Get-EffectiveConfiguration -ConfigPath $ConfigPath -ScriptDirectory $state.ScriptDirectory -NoEnvironment
                if (Save-Configuration -Path $ConfigPath -Configuration $effective.Config) {
                    Write-LogSuccess "Configuration written to '$ConfigPath' (source: $($effective.Source))." -Component 'Config'
                    $null = Set-ScriptExitCode $Script:ExitCode.Success
                }
                else {
                    $null = Set-ScriptExitCode $Script:ExitCode.Failed
                }
                return
            }

            if ($Restore) {
                $state.Mode = 'Restore'
                try {
                    $params = @{ Path = $Restore; Force = $Force }
                    if (-not [string]::IsNullOrWhiteSpace($RestoreDestination)) { $params['DestinationDirectory'] = $RestoreDestination }
                    $result = Restore-SpoolerBackupInternal @params

                    if ($state.MachineReadable) { ($result | ConvertTo-Json -Depth 6) | Write-Output }
                    else {
                        Write-ConsoleSection -Title 'BACKUP RESTORED' -Character '='
                        Write-ConsoleKeyValue -Label 'Archive' -Value $result.BackupPath
                        Write-ConsoleKeyValue -Label 'Source computer' -Value $result.ComputerName
                        Write-ConsoleKeyValue -Label 'Taken (UTC)' -Value $result.CreatedUtc
                        Write-ConsoleKeyValue -Label 'Files in manifest' -Value $result.ManifestFileCount
                        Write-ConsoleKeyValue -Label 'Files restored' -Value $result.FilesRestored -ValueColor 'Green'
                        Write-ConsoleKeyValue -Label 'Destination' -Value $result.Destination
                        Write-ConsoleKeyValue -Label 'SHA-256' -Value $result.Sha256
                        Write-Banner -Color 'Yellow' -Text @(
                            'Spool files cannot be re-queued by copying them back into PRINTERS.',
                            'Re-submit the recovered documents from the application.'
                        )
                    }
                    $null = Set-ScriptExitCode $Script:ExitCode.Success
                    if ($PassThru) { return $result }
                }
                catch {
                    Write-LogError "Restore failed: $($_.Exception.Message)" -Component 'Restore'
                    $null = Set-ScriptExitCode $Script:ExitCode.RestoreFailed
                }
                return
            }

            if ($CheckUpdate -or $Update) {
                if ($CheckUpdate -and -not $Update) {
                    try {
                        $info = Get-SpoolerCleanupUpdate -Channel $UpdateChannel -GitHubToken $GitHubToken -IncludePrerelease:$IncludePrerelease
                        if ($state.MachineReadable) {
                            ($info | ConvertTo-Json -Depth 6) | Write-Output
                        }
                        else {
                            Write-ConsoleSection -Title 'UPDATE CHECK' -Character '='
                            Write-ConsoleKeyValue -Label 'Installed' -Value $Script:ScriptVersion
                            Write-ConsoleKeyValue -Label 'Latest published' -Value $info.LatestVersion
                            Write-ConsoleKeyValue -Label 'Discovery source' -Value $info.Source
                            $availColor = if ($info.UpdateAvailable) { 'Yellow' } else { 'Green' }
                            Write-ConsoleKeyValue -Label 'Update available' -Value $info.UpdateAvailable -ValueColor $availColor
                            if ($info.TagName) { Write-ConsoleKeyValue -Label 'Tag' -Value $info.TagName }
                            if ($info.PublishedUtc) { Write-ConsoleKeyValue -Label 'Published' -Value $info.PublishedUtc }
                            if ($info.ExpectedSha256) { Write-ConsoleKeyValue -Label 'Published SHA-256' -Value $info.ExpectedSha256 }
                            if ($info.UpdateAvailable) {
                                Write-Host ''
                                Write-Host '    To install:  .\Clear-PrintSpoolerQueue.ps1 -Update -Force' -ForegroundColor Cyan
                            }
                            if ($info.ReleaseNotes) {
                                Write-Host ''
                                Write-ConsoleSection -Title 'RELEASE NOTES'
                                foreach ($line in (Split-HelpParagraph -Text $info.ReleaseNotes -Width 100)) { Write-Host "    $line" }
                            }
                            foreach ($err in @($info.Errors)) { Write-LogWarn "Discovery: $err" -Component 'Update' }
                        }
                        $code = $Script:ExitCode.Success
                        if (-not $info.LatestVersion) { $code = $Script:ExitCode.UpdateCheckFailed }
                        elseif ($info.UpdateAvailable) { $code = $Script:ExitCode.UpdateAvailable }
                        $null = Set-ScriptExitCode $code
                        if ($PassThru) { return $info }
                    }
                    catch {
                        Write-LogError "Update check failed: $($_.Exception.Message)" -Component 'Update'
                        $null = Set-ScriptExitCode $Script:ExitCode.UpdateCheckFailed
                    }
                    return
                }

                # ---- -Update
                try {
                    $params = @{
                        Channel                  = $UpdateChannel
                        Force                    = $Force
                        ApplyNow                 = $ApplyUpdateNow
                        SkipChecksumVerification = $SkipChecksumVerification
                        IncludePrerelease        = $IncludePrerelease
                        InstallDirectory         = $state.InstallDirectory
                    }
                    if (-not [string]::IsNullOrWhiteSpace($GitHubToken)) { $params['GitHubToken'] = $GitHubToken }
                    $result = Update-SpoolerCleanup @params

                    if ($state.MachineReadable) { ($result | ConvertTo-Json -Depth 6) | Write-Output }
                    else {
                        Write-ConsoleSection -Title 'UPDATE RESULT' -Character '='
                        Write-ConsoleKeyValue -Label 'From' -Value $result.FromVersion
                        Write-ConsoleKeyValue -Label 'To' -Value $result.ToVersion
                        $vColor = if ($result.VerifyResult -eq 'ChecksumMatch') { 'Green' } else { 'Yellow' }
                        Write-ConsoleKeyValue -Label 'Verification' -Value $result.VerifyResult -ValueColor $vColor
                        Write-ConsoleKeyValue -Label 'Method' -Value $result.Method
                        Write-ConsoleKeyValue -Label 'Staged' -Value $result.Staged
                        Write-ConsoleKeyValue -Label 'Applied' -Value $result.Updated -ValueColor $(if ($result.Updated) { 'Green' } else { 'Yellow' })
                        Write-ConsoleKeyValue -Label 'Rollback copy' -Value $result.BackupPath
                        Write-ConsoleKeyValue -Label 'Message' -Value $result.Message
                    }
                    $code = if ($result.Updated -or $result.Staged) { $Script:ExitCode.UpdateApplied } else { $Script:ExitCode.Success }
                    $null = Set-ScriptExitCode $code
                    if ($PassThru) { return $result }
                }
                catch {
                    Write-LogError "Update failed: $($_.Exception.Message)" -Component 'Update'
                    $null = Set-ScriptExitCode $Script:ExitCode.UpdateFailed
                }
                return
            }

            # ================================================================
            #  STEP 3 - platform, elevation, signature
            # ================================================================
            if (-not (Test-RunningOnWindows)) {
                Write-LogError 'This tool manages the Windows Print Spooler and cannot run on this platform.' -Component 'Platform'
                $null = Set-ScriptExitCode $Script:ExitCode.UnsupportedPlatform
                return
            }
            if (-not (Test-PrintSpoolerAvailable)) {
                Write-LogError ("The Print Spooler service ('{0}') is not installed on this host. Install the Print Services feature, or run against a remote target with -ComputerName." -f $Script:SpoolerServiceName) -Component 'Platform'
                $null = Set-ScriptExitCode $Script:ExitCode.UnsupportedPlatform
                return
            }
            if ($PSVersionTable.PSVersion -lt $Script:MinPSVersion) {
                Write-LogError ("PowerShell {0} is below the required {1}." -f $PSVersionTable.PSVersion, $Script:MinPSVersion) -Component 'Platform'
                $null = Set-ScriptExitCode $Script:ExitCode.UnsupportedPlatform
                return
            }

            if ($VerifySignature) {
                $self = Join-Path $state.ScriptDirectory "$($Script:ScriptName).ps1"
                $sig = Test-ScriptSignature -Path $self -ExpectedThumbprint $ExpectedSignerThumbprint
                Write-LogInfo ("Signature of '{0}': {1} ({2})" -f $self, $sig.Status, $sig.SignerSubject) -Component 'Security'
                if ($sig.Status -ne 'Valid') {
                    Write-LogError ("Signature verification failed: {0}. {1}" -f $sig.Status, $sig.Message) -Component 'Security'
                    $null = Set-ScriptExitCode $Script:ExitCode.SignatureInvalid
                    return
                }
                Write-LogSuccess 'Signature is valid.' -Component 'Security'
            }

            # ---- build the target list (parameters + pipeline + file + environment)
            $targets = New-Object System.Collections.Generic.List[string]
            foreach ($t in $state.TargetList.ToArray()) { if ($targets -notcontains $t) { $targets.Add($t) } }

            if (-not [string]::IsNullOrWhiteSpace($ComputerListFile)) {
                if (-not (Test-Path -LiteralPath $ComputerListFile -PathType Leaf)) {
                    Write-LogError "Computer list file not found: $ComputerListFile" -Component 'Targets'
                    $null = Set-ScriptExitCode $Script:ExitCode.BadArguments
                    return
                }
                foreach ($line in @(Get-Content -LiteralPath $ComputerListFile -ErrorAction SilentlyContinue)) {
                    $c = ([string]$line).Trim()
                    if (-not $c -or $c.StartsWith('#') -or $c.StartsWith(';')) { continue }
                    if ($targets -notcontains $c) { $targets.Add($c) }
                }
                Write-LogInfo "Loaded the computer list from '$ComputerListFile' ($($targets.Count) target(s) total)." -Component 'Targets'
            }

            if ($targets.Count -eq 0 -and $env:PSQ_COMPUTERS) {
                foreach ($c in ($env:PSQ_COMPUTERS -split '[,;]')) {
                    $t = $c.Trim()
                    if ($t -and ($targets -notcontains $t)) { $targets.Add($t) }
                }
            }
            if ($targets.Count -eq 0) { $targets.Add('.') }

            $touchesLocal = $false
            foreach ($t in $targets.ToArray()) {
                if (Test-IsLocalTarget -ComputerName $t) { $touchesLocal = $true; break }
            }

            if ($touchesLocal -and -not (Test-RunningAsAdministrator)) {
                if ($SelfElevate -and -not $NoElevate) {
                    $childArgs = New-ChildArgumentList -BoundParameters $PSBoundParameters
                    $scriptFile = Join-Path $state.ScriptDirectory "$($Script:ScriptName).ps1"
                    $elev = Invoke-SelfElevation -ScriptPath $scriptFile -ChildArguments $childArgs
                    if ($elev.Success -and $null -ne $elev.ExitCode) {
                        Write-LogSuccess $elev.Message -Component 'Security'
                        $null = Set-ScriptExitCode ([int]$elev.ExitCode)
                        return
                    }
                    Write-LogError ("Elevation did not succeed: {0}" -f $elev.Message) -Component 'Security'
                    $code = if ($elev.ExitCode -eq $Script:ExitCode.Cancelled) { $Script:ExitCode.Cancelled } else { $Script:ExitCode.NotAdministrator }
                    $null = Set-ScriptExitCode $code
                    return
                }
                if ($NoElevate) {
                    Write-LogError 'Administrator privileges are required for the local machine and -NoElevate was specified.' -Component 'Security'
                    $null = Set-ScriptExitCode $Script:ExitCode.NotAdministrator
                    return
                }
                if ($state.NonInteractive) {
                    Write-LogError 'Administrator privileges are required. The host is non-interactive, so elevation cannot be requested. Run this task as SYSTEM or from an elevated session.' -Component 'Security'
                    $null = Set-ScriptExitCode $Script:ExitCode.NotAdministrator
                    return
                }
                Write-Banner -Color 'Red' -Text @(
                    'Administrator privileges are required to manage the Print Spooler.',
                    'Right-click and choose "Run as administrator", re-run from an',
                    'elevated shell, or pass -SelfElevate to trigger a UAC prompt.'
                )
                $null = Set-ScriptExitCode $Script:ExitCode.NotAdministrator
                return
            }

            # ================================================================
            #  STEP 4 - configuration and parameter validation
            # ================================================================
            # Already loaded in the begin block, because the log target depends on
            # it. Validation is reported here rather than there so that an invalid
            # configuration still produces a log record explaining the failure.
            if (@($state.ConfigProblems).Count -gt 0) {
                foreach ($cfgProblem in @($state.ConfigProblems)) { Write-LogError "Configuration: $cfgProblem" -Component 'Config' }
                $null = Set-ScriptExitCode $Script:ExitCode.BadArguments
                return
            }
            $config = $state.Configuration
            $cfgNote = if ($state.ConfigPath) { " ($($state.ConfigPath))" } else { '' }
            Write-LogDebug ("Configuration source: {0}{1}" -f $state.ConfigSource, $cfgNote) -Component 'Config'

            # Explicit parameters beat configuration.
            if ($PSBoundParameters.ContainsKey('ServiceTimeoutSeconds')) { $config['Service']['TimeoutSeconds'] = $ServiceTimeoutSeconds }
            if ($PSBoundParameters.ContainsKey('SettleSeconds')) { $config['Service']['SettleSeconds'] = $SettleSeconds }
            if ($PSBoundParameters.ContainsKey('RetryCount')) { $config['Cleanup']['RetryCount'] = $RetryCount }
            if ($PSBoundParameters.ContainsKey('RetryDelayMs')) { $config['Cleanup']['RetryDelayMs'] = $RetryDelayMs }
            if ($PSBoundParameters.ContainsKey('StopDependentServices')) { $config['Service']['StopDependentServices'] = $StopDependentServices }
            if ($PSBoundParameters.ContainsKey('RestartDependentServices')) { $config['Service']['RestartDependentServices'] = $RestartDependentServices }
            if ($PSBoundParameters.ContainsKey('ThrottleLimit')) { $config['Execution']['ThrottleLimit'] = $ThrottleLimit }
            if ($PSBoundParameters.ContainsKey('LogRetentionDays')) { $config['Logging']['RetentionDays'] = $LogRetentionDays }
            if ($Backup) { $config['Cleanup']['Backup'] = $true }
            if (-not [string]::IsNullOrWhiteSpace($BackupPath)) { $config['Cleanup']['BackupPath'] = $BackupPath }
            if ($EventLog) { $config['Logging']['EventLog'] = $true }
            if ($JsonLog) { $config['Logging']['JsonLog'] = $true }
            if ($PSBoundParameters.ContainsKey('LogLevel')) {
                $config['Logging']['Level'] = $LogLevel
                $null = Set-RuntimeLogLevel -Level $LogLevel
            }

            $conflicts = Test-ParameterCombination -Bound $PSBoundParameters
            if (@($conflicts).Count -gt 0) {
                foreach ($c in @($conflicts)) { Write-LogError "Parameter conflict: $c" -Component 'Arguments' }
                Write-LogInfo 'Run with -Help for the full reference.' -Component 'Arguments'
                $null = Set-ScriptExitCode $Script:ExitCode.BadArguments
                return
            }

            # ---- normalise the human-friendly size arguments
            $largerBytes = [long]0
            $smallerBytes = [long]0
            if ($PSBoundParameters.ContainsKey('LargerThanBytes') -and $null -ne $LargerThanBytes) {
                if ($LargerThanBytes -is [string]) {
                    $parsed = Format-FileSizeArgument -Value $LargerThanBytes
                    if ($null -eq $parsed) {
                        Write-LogError "-LargerThanBytes '$LargerThanBytes' is not a valid size (try 512, 10MB, 1.5GB)." -Component 'Arguments'
                        $null = Set-ScriptExitCode $Script:ExitCode.BadArguments
                        return
                    }
                    $largerBytes = [long]$parsed
                }
                else { $largerBytes = [long]$LargerThanBytes }
            }
            if ($PSBoundParameters.ContainsKey('SmallerThanBytes') -and $null -ne $SmallerThanBytes) {
                if ($SmallerThanBytes -is [string]) {
                    $parsed = Format-FileSizeArgument -Value $SmallerThanBytes
                    if ($null -eq $parsed) {
                        Write-LogError "-SmallerThanBytes '$SmallerThanBytes' is not a valid size." -Component 'Arguments'
                        $null = Set-ScriptExitCode $Script:ExitCode.BadArguments
                        return
                    }
                    $smallerBytes = [long]$parsed
                }
                else { $smallerBytes = [long]$SmallerThanBytes }
            }
            if ($OlderThan) {
                try { $null = ConvertFrom-AnyTimeSpan -Value $OlderThan }
                catch {
                    Write-LogError $_.Exception.Message -Component 'Arguments'
                    $null = Set-ScriptExitCode $Script:ExitCode.BadArguments
                    return
                }
            }
            if ($NewerThan) {
                try { $null = ConvertFrom-AnyTimeSpan -Value $NewerThan }
                catch {
                    Write-LogError $_.Exception.Message -Component 'Arguments'
                    $null = Set-ScriptExitCode $Script:ExitCode.BadArguments
                    return
                }
            }

            # ---- log rotation and retention (cheap, and keeps the disk safe)
            if ($state.LogPath) {
                if (-not $NoLogRotation) {
                    $null = Invoke-LogRotation -Path $state.LogPath -MaxSizeMB ([double]$config['Logging']['MaxSizeMB']) -MaxFileCount ([int]$config['Logging']['MaxFileCount'])
                }
                $pruned = Invoke-LogRetention -Path $state.LogPath -RetentionDays ([int]$config['Logging']['RetentionDays'])
                if ($pruned -gt 0) { Write-LogDebug "Log retention removed $pruned old log file(s)." -Component 'Logging' }
            }

            # ---- automatic update check (quiet, cooldown-limited, never fatal)
            if (-not $NoUpdateCheck -and $config['Update']['Enabled'] -and $config['Update']['AutoCheck']) {
                $token = $GitHubToken
                if ([string]::IsNullOrWhiteSpace($token)) { $token = [string]$config['Update']['Token'] }
                if ([string]::IsNullOrWhiteSpace($token)) { $token = $env:PSQ_GITHUB_TOKEN }
                $null = Invoke-AutoUpdateCheck -InstallDirectory $state.InstallDirectory `
                    -CooldownHours ([int]$config['Update']['CooldownHours']) -Channel $UpdateChannel -Token $token -ApplyNow:$ApplyUpdateNow
            }

            # ================================================================
            #  STEP 5 - build the payload and show the plan
            # ================================================================
            $arguments = @{
                Mode                     = $state.Mode
                DryRun                   = [bool]$WhatIfPreference
                Repair                   = [bool]$Repair
                Backup                   = [bool]$config['Cleanup']['Backup']
                BackupRequired           = [bool]$BackupRequired
                BackupPath               = [string]$config['Cleanup']['BackupPath']
                SpoolDirectory           = $SpoolDirectory
                PrinterName              = $PrinterName
                ExcludePrinterName       = $ExcludePrinterName
                JobId                    = $JobId
                DocumentName             = $DocumentName
                OwnerName                = $OwnerName
                OlderThan                = $OlderThan
                NewerThan                = $NewerThan
                LargerThanBytes          = $largerBytes
                SmallerThanBytes         = $smallerBytes
                PurgeFormsCache          = [bool]($PurgeFormsCache -and $Repair)
                PurgeDriverStore         = [bool]($PurgeDriverStore -and $Repair)
                ResetSpoolDirectoryValue = [bool]($ResetSpoolDirectory -and $Repair)
            }
            $options = New-PayloadOptions -Config $config -Arguments $arguments

            $payload = $null
            try { $payload = Resolve-PayloadScriptBlock }
            catch {
                Write-LogError $_.Exception.Message -Component 'Payload'
                $null = Set-ScriptExitCode $Script:ExitCode.Failed
                return
            }

            if (-not $state.MachineReadable) {
                Write-ConsoleSection -Title 'EXECUTION PLAN'
                $targetList = (@($targets | Select-Object -First 6) -join ', ')
                if ($targets.Count -gt 6) { $targetList += ', ...' }
                Write-ConsoleKeyValue -Label 'Mode' -Value $state.Mode
                Write-ConsoleKeyValue -Label 'Targets' -Value ("{0} ({1})" -f $targets.Count, $targetList)
                Write-ConsoleKeyValue -Label 'Selection' -Value (Format-SelectionDescription -Arguments $arguments)
                $backupText = if ($arguments.Backup) { if ($arguments.BackupPath) { "yes -> $($arguments.BackupPath)" } else { 'yes -> default location' } } else { 'no' }
                Write-ConsoleKeyValue -Label 'Backup' -Value $backupText
                $depText = if ($config['Service']['StopDependentServices']) { 'stopped first, restarted after' } else { 'left alone' }
                Write-ConsoleKeyValue -Label 'Dependent services' -Value $depText
                Write-ConsoleKeyValue -Label 'Service timeout' -Value ("{0}s" -f $config['Service']['TimeoutSeconds'])
                Write-ConsoleKeyValue -Label 'Delete retries' -Value ("{0} (initial delay {1} ms, exponential)" -f $config['Cleanup']['RetryCount'], $config['Cleanup']['RetryDelayMs'])
                if ($targets.Count -gt 1) { Write-ConsoleKeyValue -Label 'Concurrency' -Value ("{0} worker(s)" -f [Math]::Max(1, [int]$config['Execution']['ThrottleLimit'])) }
                Write-ConsoleKeyValue -Label 'Log file' -Value $(if ($state.LogPath) { $state.LogPath } else { 'disabled' })
                if ($WhatIfPreference) { Write-ConsoleKeyValue -Label 'Changes' -Value 'NONE (dry run)' -ValueColor 'Cyan' }
            }

            # ================================================================
            #  STEP 6 - confirmation gate
            # ================================================================
            $selective = [bool]($PrinterName -or $ExcludePrinterName -or $JobId -or $DocumentName -or $OwnerName -or $OlderThan -or $NewerThan -or ($largerBytes -gt 0) -or ($smallerBytes -gt 0))
            if (-not $WhatIfPreference -and -not $Diagnostic -and -not $Force) {
                $confirmTarget = if ($targets.Count -eq 1) { "the print queue on '$($targets[0])'" } else { "the print queues on $($targets.Count) computers" }
                $scope = if ($selective) { 'the SELECTED' } else { 'ALL' }
                $confirmAction = "Stop the Print Spooler, delete $scope spool files, then restart the service"

                $proceed = Confirm-DestructiveAction -Target $confirmTarget -Action $confirmAction -Force:$Force `
                    -DefaultAnswer 'Y' -TimeoutSeconds ([int]$config['Execution']['ConfirmTimeoutSeconds'])
                if (-not $proceed) {
                    Write-LogWarn 'Operation cancelled by the operator.' -Component 'Confirm'
                    $null = Set-ScriptExitCode $Script:ExitCode.Cancelled
                    return
                }
            }

            # ================================================================
            #  STEP 7 - watchdog
            # ================================================================
            if ($Watchdog -and $WatchdogMinutes -gt 0) {
                $state.WatchdogJob = Start-WatchdogInternal -Minutes $WatchdogMinutes -InvocationId $state.InvocationId -Targets $targets.ToArray()
                Write-LogInfo ("Watchdog armed: the run will be terminated and the Spooler restarted if it exceeds {0} minute(s)." -f $WatchdogMinutes) -Component 'Watchdog'
            }

            # ================================================================
            #  STEP 8 - dispatch
            # ================================================================
            $dispatchParams = @{
                Targets             = $targets.ToArray()
                Options             = $options
                Payload             = $payload
                ThrottleLimit       = [Math]::Max(1, [int]$config['Execution']['ThrottleLimit'])
                PreFlightProbe      = [bool]$PreFlightProbe
                ProbeTimeoutSeconds = [int]$config['Remote']['ProbeTimeoutSeconds']
                UseSsl              = $UseSsl
            }
            if ($Credential) { $dispatchParams['Credential'] = $Credential }

            $results = @(Invoke-CleanupFleet @dispatchParams)

            # ---- optional proof that printing really works again
            if ($SelfTestPrint -and -not $WhatIfPreference -and -not $Diagnostic) {
                foreach ($r in $results) {
                    if (($r.Status -in @('Success', 'SuccessWithWarnings')) -and $r.ServiceIsRunning) {
                        $probe = Invoke-SelfTestPrintInternal -ComputerName $r.ComputerName -Credential $Credential
                        $r | Add-Member -NotePropertyName SelfTestPrint -NotePropertyValue $probe -Force
                        if ($probe.Success) {
                            Write-LogSuccess ("Self-test print job on '{0}' was accepted and cancelled - the queue is functional." -f $r.ComputerName) -Component 'SelfTest' -Target $r.ComputerName
                        }
                        else {
                            Write-LogWarn ("Self-test print on '{0}' failed: {1}" -f $r.ComputerName, $probe.Message) -Component 'SelfTest' -Target $r.ComputerName
                            Add-ResultWarning -Result $r -Message "Self-test print failed: $($probe.Message)"
                        }
                    }
                }
            }

            # ---- move the backups off-box if asked
            if (-not [string]::IsNullOrWhiteSpace($ArchivePath) -and (Test-Path -LiteralPath $ArchivePath -PathType Container)) {
                foreach ($r in $results) {
                    if ($r.BackupPath -and (Test-Path -LiteralPath $r.BackupPath -PathType Leaf)) {
                        $leaf = [IO.Path]::GetFileName($r.BackupPath)
                        try {
                            Move-Item -LiteralPath $r.BackupPath -Destination $ArchivePath -Force -ErrorAction Stop
                            Write-LogInfo ("Backup archive moved to '{0}'." -f (Join-Path $ArchivePath $leaf)) -Component 'Backup' -Target $r.ComputerName
                            $r.BackupPath = Join-Path $ArchivePath $leaf
                        }
                        catch {
                            Write-LogWarn ("Could not archive the backup for '{0}': {1}" -f $r.ComputerName, $_.Exception.Message) -Component 'Backup'
                        }
                    }
                }
            }

            # ---- backup retention
            if ($config['Cleanup']['Backup']) {
                $bkRoot = $null
                if ($config['Cleanup']['BackupPath']) { $bkRoot = [string]$config['Cleanup']['BackupPath'] }
                elseif ($env:ProgramData) { $bkRoot = Join-Path (Join-Path $env:ProgramData $Script:ProductFamily) 'Backups' }
                $pruned = Invoke-BackupRetention -BackupPath $bkRoot -KeepDays ([int]$config['Cleanup']['BackupKeepDays'])
                if ($pruned -gt 0) { Write-LogDebug "Backup retention removed $pruned old bundle(s)." -Component 'Backup' }
            }

            # ================================================================
            #  STEP 9 - aggregate, report, export, exit code
            # ================================================================
            $sw.Stop()
            $summary = New-RunSummary -Results $results -Arguments $arguments -LogPath $state.LogPath `
                -ConfigSource $state.ConfigSource -Duration $sw.Elapsed
            $exitCode = Get-ExitCodeForSummary -Summary $summary

            # A service that could not be stopped or started outranks "partial success".
            $serviceFailures = @($results | Where-Object { $_.Status -eq 'ServiceFailure' }).Count
            if (($serviceFailures -gt 0) -and ($exitCode -lt $Script:ExitCode.ServiceFailure)) {
                $exitCode = $Script:ExitCode.ServiceFailure
            }
            if ($Diagnostic -and ($exitCode -eq $Script:ExitCode.Success)) { $exitCode = $Script:ExitCode.DiagnosticsOnly }

            $summary.ExitCode = $exitCode
            $state.Summary = $summary
            $state.Results = $results

            Write-LogInfo ("Run finished. Status={0}; targets={1}; jobsRemoved={2}; filesDeleted={3}; reclaimed={4}; exitCode={5}." -f `
                $summary.OverallStatus, $summary.TargetsAttempted, $summary.JobsRemoved, $summary.FilesDeleted, `
                (Format-ByteSize $summary.BytesDeleted), $exitCode) -Component 'Session'

            if ($Diagnostic) { Write-DiagnosticReport -Summary $summary }
            Write-RunSummaryToConsole -Summary $summary

            if ($config['Logging']['EventLog']) {
                $evtType = switch ($summary.OverallStatus) {
                    'Success' { 'SuccessAudit' }
                    'NothingToDo' { 'Information' }
                    'DryRun' { 'Information' }
                    'Diagnostic' { 'Information' }
                    'SuccessWithWarnings' { 'Warning' }
                    'PartialSuccess' { 'Warning' }
                    default { 'Error' }
                }
                $evtId = switch ($evtType) {
                    'SuccessAudit' { $Script:EventLogIdSuccess }
                    'Warning' { $Script:EventLogIdWarning }
                    'Error' { $Script:EventLogIdError }
                    default { $Script:EventLogIdInfo }
                }
                $msg = ("{0} finished with status {1}. Targets: {2} attempted, {3} succeeded, {4} failed. Jobs removed: {5}. Files deleted: {6}. Bytes reclaimed: {7}. Exit code: {8}." -f `
                    $Script:ScriptName, $summary.OverallStatus, $summary.TargetsAttempted, $summary.TargetsSucceeded, `
                    $summary.TargetsFailed, $summary.JobsRemoved, $summary.FilesDeleted, $summary.BytesDeleted, $exitCode)
                $null = Write-EventLogRecord -Message $msg -EntryType $evtType -EventId $evtId
            }

            # Blanks are filtered before use. A blank entry can come from an unset or
            # empty -ExportFormat value, and it is falsy in PowerShell: forwarded as a
            # format name it would be parsed as text and silently write a report nobody
            # asked for - and, before the binder was relaxed, it aborted the whole run.
            # The configured list gets the same treatment.
            $exportFormats = @($ExportFormat | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            if ($exportFormats.Count -eq 0) {
                $exportFormats = @($config['Execution']['ExportFormat'] | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            }
            if ($exportFormats.Count -gt 0) {
                $exportTarget = $ExportPath
                if ([string]::IsNullOrWhiteSpace($exportTarget)) { $exportTarget = [string]$config['Execution']['ExportPath'] }
                $files = @(Export-RunReport -Summary $summary -Format $exportFormats -Path $exportTarget)
                if (-not $state.MachineReadable) {
                    foreach ($f in $files) { Write-ConsoleKeyValue -Label 'Report' -Value $f }
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($DiagnosticsBundlePath)) {
                $bundle = Write-DiagnosticsBundle -DestinationDirectory $DiagnosticsBundlePath -Summary $summary -Configuration $config
                Write-LogSuccess "Diagnostics bundle written: $bundle" -Component 'Diagnostics'
                if (-not $state.MachineReadable) { Write-ConsoleKeyValue -Label 'Diagnostics bundle' -Value $bundle }
            }

            if ($state.MachineReadable) {
                # Exactly one JSON document on stdout, and nothing else.
                (ConvertTo-JsonReport -Summary $summary) | Write-Output
            }

            if ($state.WatchdogJob) { Stop-WatchdogInternal -Job $state.WatchdogJob }

            $null = Set-ScriptExitCode $exitCode
            if ($PassThru) { return $summary }
        }
        catch {
            $message = $_.Exception.Message
            Write-LogCritical "Unhandled error: $message" -Component 'Session'
            if ($_.ScriptStackTrace) { Write-LogDebug $_.ScriptStackTrace -Component 'Session' }
            if (-not $state.MachineReadable -and -not $state.Quiet) {
                Write-Banner -Color 'Red' -Text @('The run failed with an unhandled error.', $message)
                Write-Host ''
                Write-Host '    A diagnostics bundle helps enormously when reporting this:' -ForegroundColor Yellow
                Write-Host '        .\Clear-PrintSpoolerQueue.ps1 -Diagnostic -DiagnosticsBundlePath C:\Temp' -ForegroundColor Yellow
            }
            if ($state.WatchdogJob) {
                try { Stop-WatchdogInternal -Job $state.WatchdogJob }
                catch { /* teardown must never mask the original error */ }
            }
            $null = Set-ScriptExitCode $Script:ExitCode.Failed
        }
        finally {
            Show-CleanupProgress -Activity 'Print Spooler Cleanup' -Completed
        }
    }
}

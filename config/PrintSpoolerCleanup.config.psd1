# ===========================================================================
#  Print Spooler Queue Cleanup - configuration template  (v4.0.1)
# ===========================================================================
#
#  Copy this file next to Clear-PrintSpoolerQueue.ps1 (or to
#  %ProgramData%\PrintSpoolerCleanup\) and edit the values you need.
#  Anything you delete falls back to the built-in default.
#
#  SECURITY MODEL
#    This file is parsed with the PowerShell AST and evaluated with
#    SafeGetValue(). Command invocations, pipelines and variable assignments
#    are REJECTED, so a config that arrives from a file share or a Git
#    repository cannot execute code. Files larger than 1 MB are refused.
#
#  PRECEDENCE (highest wins)
#    1. Command-line parameters
#    2. PSQ_* environment variables
#    3. -ConfigPath <file>
#    4. This file, auto-discovered in:
#         <script dir> -> <current dir>
#         -> %ProgramData%\PrintSpoolerCleanup
#         -> %LocalAppData%\PrintSpoolerCleanup
#    5. Built-in defaults
#
#  VALIDATION
#    Every value is range-checked. Unknown sections or keys, and out-of-range
#    numbers, abort the run with exit code 3 rather than being ignored - a
#    silently ignored setting is how a site loses its backups.
#
#  Regenerate this file at any time with:
#      .\Clear-PrintSpoolerQueue.ps1 -SaveConfig -ConfigPath .\PrintSpoolerCleanup.config.psd1
#
#  Author : Mikhail Deynekin <Mikhail@Deynekin.com>
#  Source : https://github.com/paulmann/Print-Spooler-Queue-Cleanup
#  License: GPL-3.0-or-later
# ===========================================================================

@{

    # -----------------------------------------------------------------------
    #  Service orchestration
    # -----------------------------------------------------------------------
    Service = @{

        # Seconds to wait for a service to reach the desired state.
        # A print server with a large queue and a slow disk can need more than
        # the classic 30 s; 60 s is the production-tested default.
        TimeoutSeconds           = 60

        # Pause after the Spooler stops, so open handles are released before
        # the delete step. 0 makes locked-file failures much more likely.
        SettleSeconds            = 2

        # Stop Fax / PrintNotify / PrintWorkflowUserSvc_* BEFORE the Spooler.
        # This is the single setting that most often turns "the stop hung"
        # into a clean run. Only services that are actually running are
        # restarted afterwards, so a disabled Fax service stays disabled.
        StopDependentServices    = $true
        RestartDependentServices = $true

        # Escalate to terminating the service process when a service refuses to
        # leave STOP_PENDING. Safe for the Spooler (the SCM restarts it), and
        # the only reliable way out of a PrintNotify hang.
        ForceStopFallback        = $true

        # Prefix-matched where noted: real names are PrintWorkflowUserSvc_1a2b3c.
        DependentServices        = @('PrintWorkflowUserSvc', 'Fax', 'PrintNotify')

        # Not normally changed. Present so an unusual SKU can be supported.
        Name                     = 'Spooler'
    }

    # -----------------------------------------------------------------------
    #  Cleanup behaviour
    # -----------------------------------------------------------------------
    Cleanup = @{

        # The only extensions ever deleted. NEVER add a wildcard: the tool
        # refuses to run if this list contains '*' or '.*'.
        IncludeExtensions        = @('.SPL', '.SHD')

        # Also remove legacy driver leftovers (.SPF, ~DF*). Off by default
        # because those files are driver-specific, not spooler-owned.
        PurgeLegacyExtensions    = $false

        # Never recurse. Subdirectories of PRINTERS are not queues.
        Recurse                  = $false

        # Archive every file before deleting it. Turn this ON in production;
        # it is the difference between a fix and an incident.
        Backup                   = $true
        BackupPath               = $null      # $null = %ProgramData%\PrintSpoolerCleanup\Backups
        BackupKeepDays           = 30

        # Delete retries with exponential backoff (500 ms -> 1 s -> 2 s).
        # A file held by spoolsv.exe usually releases within a few hundred ms.
        RetryCount               = 3
        RetryDelayMs             = 500

        # Wait until the queue stops changing before stopping the service, so
        # an actively draining queue is not interrupted mid-job. 0 = do not wait.
        WaitForQuiescenceSeconds = 0

        # After the restart, re-scan and report if files reappeared. A non-zero
        # count means a driver or port monitor is re-submitting them.
        VerifyAfterRestart       = $true

        # Default selection filter when -PrinterName is not given on the
        # command line. Empty = every queued job.
        PrinterName              = @()
        ExcludePrinterName       = @()

        # Default age filter, e.g. '2h'. $null = no age filter.
        OlderThan                = $null
    }

    # -----------------------------------------------------------------------
    #  Logging
    # -----------------------------------------------------------------------
    Logging = @{

        # TRACE DEBUG INFO WARN ERROR CRITICAL
        Level                    = 'INFO'

        # $null = first writable location of:
        #   %ProgramData%\PrintSpoolerCleanup\Logs
        #   %LocalAppData%\PrintSpoolerCleanup\Logs
        #   <script dir>\Logs
        #   %TEMP%\PrintSpoolerCleanup\Logs
        Path                     = $null

        RetentionDays            = 30
        MaxSizeMB                = 10
        MaxFileCount             = 9

        # JSON Lines instead of pipe-delimited text, for log shippers.
        JsonLog                  = $false

        # Also write to the Windows Event Log (source: PrintSpoolerCleanup).
        # Creating the source needs elevation; if that fails, event logging is
        # disabled for the run and the file log continues unaffected.
        EventLog                 = $false

        Timestamps               = $true

        # $true gives each run its own file; $false appends to one rotating file.
        # Appending is usually what you want on a scheduled task.
        SessionStampedFileName   = $false
    }

    # -----------------------------------------------------------------------
    #  Execution
    # -----------------------------------------------------------------------
    Execution = @{

        # Concurrent remote workers. 1 = sequential (the safe default, because a
        # spooler restart is disruptive). Values above 1 require PowerShell 7
        # for ForEach-Object -Parallel; on Windows PowerShell 5.1 an equivalent
        # runspace pool is used automatically. Maximum 64.
        ThrottleLimit            = 1

        Force                    = $false
        NonInteractive           = $false

        # Seconds to wait for a Y/N answer before applying the default. 0 = wait
        # forever. Set to e.g. 30 on a technician laptop so an unanswered prompt
        # cannot block a maintenance window.
        ConfirmTimeoutSeconds    = 0

        # 0 disables the watchdog. With a positive value, a run that exceeds the
        # limit is terminated and the Spooler is restarted, so a wedged host is
        # never left without printing.
        WatchdogMinutes          = 0

        SelfElevate              = $false
        NoColor                  = $false
        Ascii                    = $false
        Quiet                    = $false
        NoProgress               = $false

        # stdout becomes exactly one JSON document and nothing else.
        MachineReadableJson      = $false

        # Report formats written on every run: Json, Csv, Html, Text.
        ExportFormat             = @()
        ExportPath               = $null
    }

    # -----------------------------------------------------------------------
    #  Updates
    # -----------------------------------------------------------------------
    Update = @{

        Enabled                  = $true

        # Quiet, cooldown-limited check at start-up. Never blocks or fails a run.
        AutoCheck                = $true
        CooldownHours            = 24

        # latest | stable | prerelease
        Channel                  = 'latest'
        IncludePrerelease        = $false

        # Verify the download against the release's checksums.txt. Turning this
        # off means installing unverified bytes - do not do it.
        VerifyChecksum           = $true

        Repository               = 'paulmann/Print-Spooler-Queue-Cleanup'
        Branch                   = 'main'

        # Prefer the PSQ_GITHUB_TOKEN environment variable over putting a token
        # in a file. A token here raises the anonymous 60 req/h API limit to 5000.
        Token                    = $null
    }

    # -----------------------------------------------------------------------
    #  Diagnostics
    # -----------------------------------------------------------------------
    Diagnostics = @{
        IncludeEventLog          = $true
        EventLogHours            = 72
        EventLogMaxEntries       = 200
        IncludeDriverStore       = $true
        IncludePortMonitor       = $true
        IncludeRegistry          = $true
    }

    # -----------------------------------------------------------------------
    #  Remote execution
    # -----------------------------------------------------------------------
    Remote = @{

        # Probe WinRM on every remote target before doing real work, so offline
        # hosts fail in ~1 s instead of ~40 s each and never delay healthy ones.
        PreFlightProbe           = $true
        ProbeTimeoutSeconds      = 10

        SessionReuse             = $true
        Authentication           = 'Default'
        UseSsl                   = $false
    }
}

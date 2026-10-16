# ============================================================================
#  PUBLIC CMDLETS
# ----------------------------------------------------------------------------
#  These are the functions exported by PrintSpoolerCleanup.psd1 and available
#  after dot-sourcing Clear-PrintSpoolerQueue.ps1. They are thin, documented
#  wrappers over the internal engine, so the same logic serves all three surfaces:
#  the standalone .ps1, the .bat launcher and the importable module.
# ============================================================================

function Get-SpoolerCleanupHelp {
    <#
    .SYNOPSIS
        Comprehensive built-in reference for Print Spooler Queue Cleanup.
    .DESCRIPTION
        Renders the full operator reference: usage, modes, every parameter with its
        live type and description, worked examples, the exit-code contract, the
        configuration schema, environment variables, log and report formats, remote
        prerequisites, scheduling, update semantics, the safety model and a
        troubleshooting matrix.

        The parameter table is generated from the actual parameter metadata of
        Clear-PrintSpoolerQueue, so it cannot drift away from the code.
    .PARAMETER Section
        Render a single section instead of the whole document.
    .PARAMETER Search
        Keep only the blocks that contain this term.
    .PARAMETER Width
        Wrap width (default: current console width, capped at 140).
    .PARAMETER Sections
        List the available section names and exit.
    .EXAMPLE
        Get-SpoolerCleanupHelp
        Shows the entire reference.
    .EXAMPLE
        Get-SpoolerCleanupHelp -Section ExitCodes
    .EXAMPLE
        Get-SpoolerCleanupHelp -Search backup
        Shows every help block that mentions backups.
    .EXAMPLE
        .\Clear-PrintSpoolerQueue.ps1 -HelpSection Updates
        The same content through the script's own switch.
    .OUTPUTS
        System.String
    .NOTES
        Author : Mikhail Deynekin <Mikhail@Deynekin.com>
        Site   : https://Deynekin.com
        Repo   : https://github.com/paulmann/Print-Spooler-Queue-Cleanup
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Position = 0)]
        [ValidateSet('All', 'Usage', 'QuickStart', 'Modes', 'Parameters', 'Examples',
            'ExitCodes', 'Configuration', 'Environment', 'Logging', 'Reports',
            'Remote', 'Scheduling', 'Updates', 'Safety', 'Troubleshooting',
            'Diagnostics', 'Version', 'Credits')]
        [string]$Section = 'All',

        [Parameter(Position = 1)]
        [AllowNull()][AllowEmptyString()][string]$Search,

        [Parameter()]
        [ValidateRange(0, 200)]
        [int]$Width = 0,

        [Parameter()]
        [switch]$Sections
    )

    if ($Sections) { return ((Get-HelpSectionNames) -join [Environment]::NewLine) }
    return Get-HelpText -Section $Section -Search $Search -Width $Width -IncludeParameterMetadata
}

function Get-SpoolerCleanupVersion {
    <#
    .SYNOPSIS
        Version, build, host and licensing facts for the installed tool.
    .DESCRIPTION
        Returns a single object; with -Detailed it also reports OS, PowerShell,
        elevation, console capability and file integrity. This is what an asset
        inventory or a compliance scan should collect.
    .PARAMETER Detailed
        Include environment and file-integrity facts.
    .EXAMPLE
        Get-SpoolerCleanupVersion
    .EXAMPLE
        Get-SpoolerCleanupVersion -Detailed | ConvertTo-Json -Depth 5
    .OUTPUTS
        PSCustomObject
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([switch]$Detailed)

    $null = Get-RuntimeState
    $info = [ordered]@{
        Tool                  = $Script:ScriptName
        ProductFamily         = $Script:ProductFamily
        Version               = $Script:ScriptVersion
        Build                 = $Script:ScriptBuild
        Author                = $Script:ScriptAuthor
        Email                 = $Script:ScriptAuthorEmail
        Site                  = $Script:ScriptAuthorSite
        Repository            = $Script:ScriptRepoUrl
        Wiki                  = $Script:ScriptWikiUrl
        License               = $Script:ScriptLicense
        MinimumPowerShell     = $Script:MinPSVersion.ToString()
        RecommendedPowerShell = $Script:RecommendedPSVersion.ToString()
        InstallDirectory      = $Script:RuntimeState.InstallDirectory
    }

    if ($Detailed) {
        $info['HostPowerShell'] = $PSVersionTable.PSVersion.ToString()
        $info['PSEdition'] = $PSVersionTable.PSEdition
        $info['OperatingSystem'] = Get-WindowsSku
        $info['IsWindows'] = Test-RunningOnWindows
        $info['IsAdministrator'] = Test-RunningAsAdministrator
        $info['Context'] = Get-CurrentUserContext
        $info['Console'] = Get-ConsoleCapability
        $info['LogPath'] = $Script:RuntimeState.LogPath

        $files = New-Object System.Collections.Generic.List[object]
        foreach ($f in @('Clear-PrintSpoolerQueue.ps1', 'Clear-PrintSpoolerQueue.bat')) {
            $p = Join-Path $Script:RuntimeState.InstallDirectory $f
            if (Test-Path -LiteralPath $p) { $files.Add((Get-FileIntegrityRecord -Path $p)) }
        }
        $info['Files'] = $files.ToArray()
    }
    return [pscustomobject]$info
}

function Test-PrintSpoolerHealth {
    <#
    .SYNOPSIS
        Non-destructive health probe for the print subsystem of one or more computers.
    .DESCRIPTION
        Answers "is printing healthy?" in a form a monitoring system can consume.
        The Spooler is never stopped and no file is touched.

        Health is computed from weighted findings:
          * Spooler service state and start mode
          * spoolsv.exe responsiveness and working set
          * dependent-service state (PrintNotify, Fax)
          * queue depth and the age of the oldest job
          * total spool size
          * printers flagged offline or in error
          * recent service-crash events (SCM 7031/7034)
    .PARAMETER ComputerName
        Targets. Defaults to the local machine. Accepts pipeline input.
    .PARAMETER Credential
        Credential for remote targets.
    .PARAMETER MaxQueueAgeHours
        Queue age above which a host is degraded. Default 24.
    .PARAMETER ThrottleLimit
        Concurrent remote probes (PowerShell 7).
    .EXAMPLE
        Test-PrintSpoolerHealth
    .EXAMPLE
        Test-PrintSpoolerHealth -ComputerName SRV-PRINT01, SRV-PRINT02 | Format-Table ComputerName, Health, Score, Findings
    .EXAMPLE
        Get-Content hosts.txt | Test-PrintSpoolerHealth | Where-Object Health -ne 'Healthy' | Export-Csv unhealthy.csv
    .OUTPUTS
        PSCustomObject { ComputerName, Health, Score, SpoolerState, QueueDepth,
        OldestJobHours, SpoolBytes, SpoolDirectory, Findings[], CheckedUtc }
        Health is one of: Healthy, Degraded, Unhealthy, Unknown.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
        [Alias('Computer', 'CN', 'HostName')]
        [AllowNull()][AllowEmptyCollection()][string[]]$ComputerName,

        [Parameter()]
        [Alias('RunAs')]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter()]
        [ValidateRange(1, 720)]
        [int]$MaxQueueAgeHours = 24,

        [Parameter()]
        [ValidateRange(1, 64)]
        [int]$ThrottleLimit = 1
    )

    begin {
        $null = Get-RuntimeState
        $targets = New-Object System.Collections.Generic.List[string]
    }
    process {
        foreach ($c in @($ComputerName)) {
            if ([string]::IsNullOrWhiteSpace($c)) { continue }
            $t = $c.Trim()
            if ($targets -notcontains $t) { $targets.Add($t) }
        }
    }
    end {
        if ($targets.Count -eq 0) { $targets.Add('.') }

        # The probe runs ON the target, so it must be self-contained and use only
        # cmdlets that exist in Windows PowerShell 5.1.
        $probe = {
            param($MaxAgeHours)
            $findings = New-Object System.Collections.Generic.List[string]
            $score = 100
            $spoolerState = 'Unknown'
            $queueDepth = 0
            $oldestHours = 0.0
            $spoolBytes = [long]0

            try {
                $svc = Get-Service -Name 'Spooler' -ErrorAction Stop
                $spoolerState = [string]$svc.Status
                if ($svc.Status -ne 'Running') {
                    $findings.Add("The Print Spooler is $($svc.Status).")
                    $score -= 60
                }
                if ($svc.StartType -eq 'Disabled') {
                    $findings.Add('The Print Spooler start type is Disabled.')
                    $score -= 15
                }
            }
            catch {
                $findings.Add("The Print Spooler service is not available: $($_.Exception.Message)")
                $score -= 70
            }

            foreach ($dep in @('PrintNotify', 'Fax')) {
                $d = Get-Service -Name $dep -ErrorAction SilentlyContinue
                if ($d -and ($d.StartType -ne 'Disabled') -and ($d.Status -ne 'Running')) {
                    $findings.Add("$dep is $($d.Status) (start type $($d.StartType)).")
                    $score -= 10
                }
            }

            $proc = Get-Process -Name spoolsv -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($proc) {
                if (-not $proc.Responding) {
                    $findings.Add('spoolsv.exe is not responding.')
                    $score -= 35
                }
                $ws = [Math]::Round($proc.WorkingSet64 / 1MB, 0)
                if ($ws -gt 512) {
                    $findings.Add("The spoolsv.exe working set is $ws MB (possible leak).")
                    $score -= 15
                }
            }

            $systemRoot = $env:SystemRoot
            if ([string]::IsNullOrWhiteSpace($systemRoot)) { $systemRoot = 'C:\Windows' }
            $dir = Join-Path $systemRoot 'System32\spool\PRINTERS'
            try {
                $reg = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers' -Name 'DefaultSpoolDirectory' -ErrorAction SilentlyContinue
                if ($reg -and $reg.DefaultSpoolDirectory) { $dir = [string]$reg.DefaultSpoolDirectory }
            }
            catch { /* the OS default is a safe fallback */ }

            if (Test-Path -LiteralPath $dir) {
                $files = @(Get-ChildItem -LiteralPath $dir -File -Force -ErrorAction SilentlyContinue |
                    Where-Object { $_.Extension -ieq '.SPL' -or $_.Extension -ieq '.SHD' })
                $queueDepth = $files.Count
                $spoolBytes = [long]0
                foreach ($sb in $files) { $spoolBytes += [long]$sb.Length }
                if ($files.Count -gt 0) {
                    $oldest = @($files | Sort-Object LastWriteTimeUtc | Select-Object -First 1)[0]
                    $oldestHours = [Math]::Round(((Get-Date) - $oldest.LastWriteTime).TotalHours, 1)
                    if ($oldestHours -gt $MaxAgeHours) {
                        $findings.Add("The oldest queued job is $oldestHours h old ($($oldest.Name)).")
                        $score -= 25
                    }
                    if ($files.Count -gt 50) {
                        $findings.Add("Queue depth is $($files.Count) files.")
                        $score -= 15
                    }
                }
                if ($spoolBytes -gt 1GB) {
                    $findings.Add("The spool directory holds $([Math]::Round($spoolBytes / 1GB, 2)) GB.")
                    $score -= 10
                }
            }
            else {
                $findings.Add("Spool directory not found: $dir")
                $score -= 20
            }

            try {
                $offline = @(Get-CimInstance -ClassName Win32_Printer -ErrorAction SilentlyContinue |
                    Where-Object { $_.WorkOffline -or ($_.PrinterStatus -in @(3, 4, 5)) })
                foreach ($p in $offline) {
                    $findings.Add("Printer '$($p.Name)' is offline or in an error state (status $($p.PrinterStatus)).")
                    $score -= 8
                }
            }
            catch { /* a host without the PrintService CIM provider simply reports less */ }

            try {
                $crashes = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; Id = 7031, 7034; StartTime = (Get-Date).AddHours(-72) } -MaxEvents 50 -ErrorAction SilentlyContinue |
                    Where-Object { $_.Message -match 'Spooler|Print' })
                if ($crashes.Count -gt 0) {
                    $findings.Add("$($crashes.Count) spooler crash/terminate event(s) in the last 72 h.")
                    $score -= 20
                }
            }
            catch { /* event-log access is optional */ }

            $score = [Math]::Max(0, [Math]::Min(100, $score))
            $health = if ($score -ge 85) { 'Healthy' } elseif ($score -ge 60) { 'Degraded' } else { 'Unhealthy' }
            if ($spoolerState -eq 'Unknown') { $health = 'Unknown' }

            [pscustomobject]@{
                ComputerName   = $env:COMPUTERNAME
                Health         = $health
                Score          = $score
                SpoolerState   = $spoolerState
                QueueDepth     = $queueDepth
                OldestJobHours = $oldestHours
                SpoolBytes     = $spoolBytes
                SpoolDirectory = $dir
                Findings       = $findings.ToArray()
                CheckedUtc     = [DateTime]::UtcNow.ToString('o')
            }
        }

        $newUnknown = {
            param($computer, $message)
            [pscustomobject]@{
                ComputerName   = $computer
                Health         = 'Unknown'
                Score          = 0
                SpoolerState   = 'Unknown'
                QueueDepth     = 0
                OldestJobHours = 0
                SpoolBytes     = [long]0
                SpoolDirectory = $null
                Findings       = @($message)
                CheckedUtc     = [DateTime]::UtcNow.ToString('o')
            }
        }

        $results = New-Object System.Collections.Generic.List[object]
        $remote = New-Object System.Collections.Generic.List[string]

        foreach ($t in $targets.ToArray()) {
            if (Test-IsLocalTarget -ComputerName $t) {
                try { $results.Add((& $probe $MaxQueueAgeHours)) }
                catch { $results.Add((& $newUnknown $t "Local probe failed: $($_.Exception.Message)")) }
            }
            else { $remote.Add($t) }
        }

        if ($remote.Count -gt 0) {
            $out = $null
            if (($ThrottleLimit -gt 1) -and (Test-PowerShellSevenFeatures)) {
                $out = $remote.ToArray() | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
                    $c = $_
                    $sb = $using:probe
                    $ma = $using:MaxQueueAgeHours
                    $cr = $using:Credential
                    try {
                        $p = @{ ComputerName = $c; ScriptBlock = $sb; ArgumentList = @($ma); ErrorAction = 'Stop' }
                        if ($cr) { $p['Credential'] = $cr }
                        [pscustomobject]@{ Computer = $c; Result = (Invoke-Command @p); Error = $null }
                    }
                    catch {
                        [pscustomobject]@{ Computer = $c; Result = $null; Error = $_.Exception.Message }
                    }
                }
            }
            else {
                $out = foreach ($c in $remote.ToArray()) {
                    try {
                        $p = @{ ComputerName = $c; ScriptBlock = $probe; ArgumentList = @($MaxQueueAgeHours); ErrorAction = 'Stop' }
                        if ($Credential) { $p['Credential'] = $Credential }
                        [pscustomobject]@{ Computer = $c; Result = (Invoke-Command @p); Error = $null }
                    }
                    catch {
                        [pscustomobject]@{ Computer = $c; Result = $null; Error = $_.Exception.Message }
                    }
                }
            }

            foreach ($o in @($out)) {
                if ($o.Error) { $results.Add((& $newUnknown $o.Computer "Probe failed: $($o.Error)")) }
                else { $results.Add($o.Result) }
            }
        }

        return $results.ToArray()
    }
}

function Get-PrintSpoolerDiagnostic {
    <#
    .SYNOPSIS
        Read-only deep inspection of the print subsystem (the -Diagnostic mode
        exposed as a cmdlet).
    .DESCRIPTION
        Collects services, spooler processes, printers, drivers, ports, the spool
        directory contents, the queued-job table, relevant registry settings and
        the related event-log entries. Never stops a service, never deletes a file.
    .PARAMETER ComputerName
        Targets. Defaults to the local machine.
    .PARAMETER ExportFormat
        Report formats: Json, Csv, Html, Text. Default: console only.
    .PARAMETER ExportPath
        Directory or file to write the report to.
    .PARAMETER EventLogHours
        How many hours of event-log history to collect. Default 72.
    .EXAMPLE
        Get-PrintSpoolerDiagnostic
    .EXAMPLE
        Get-PrintSpoolerDiagnostic -ComputerName SRV-PRINT01 -ExportFormat Html -ExportPath C:\Temp
    .EXAMPLE
        Get-PrintSpoolerDiagnostic | Select-Object -ExpandProperty Results | Select-Object -ExpandProperty Jobs
    .OUTPUTS
        The same run-summary object that Clear-PrintSpoolerQueue -PassThru returns.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
        [Alias('Computer', 'CN')]
        [AllowNull()][AllowEmptyCollection()][string[]]$ComputerName,

        [Parameter()]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter()]
        [ValidateSet('Json', 'Csv', 'Html', 'Text')]
        [AllowNull()][AllowEmptyCollection()][string[]]$ExportFormat,

        [Parameter()]
        [AllowNull()][AllowEmptyString()][string]$ExportPath,

        [Parameter()]
        [ValidateRange(1, 2160)]
        [int]$EventLogHours = 72
    )

    begin { $targets = New-Object System.Collections.Generic.List[string] }
    process {
        foreach ($c in @($ComputerName)) {
            if ([string]::IsNullOrWhiteSpace($c)) { continue }
            $t = $c.Trim()
            if ($targets -notcontains $t) { $targets.Add($t) }
        }
    }
    end {
        if ($targets.Count -eq 0) { $targets.Add('.') }
        $params = @{
            ComputerName  = $targets.ToArray()
            Diagnostic    = $true
            Force         = $true
            PassThru      = $true
            EventLogHours = $EventLogHours
        }
        if ($Credential) { $params['Credential'] = $Credential }
        if ($ExportFormat -and @($ExportFormat).Count -gt 0) { $params['ExportFormat'] = $ExportFormat }
        if (-not [string]::IsNullOrWhiteSpace($ExportPath)) { $params['ExportPath'] = $ExportPath }
        return Clear-PrintSpoolerQueue @params
    }
}

function Get-SpoolerCleanupUpdate {
    <#
    .SYNOPSIS
        Reports whether a newer version is published on GitHub.
    .DESCRIPTION
        Tries three independent discovery strategies - releases/latest, tags, and
        the raw branch file - so it keeps working when the REST API is rate-limited
        or when the repository publishes no formal releases.
    .PARAMETER Channel
        latest (default), stable (no prereleases) or prerelease.
    .PARAMETER GitHubToken
        Optional token. Prefer the PSQ_GITHUB_TOKEN environment variable so the
        secret never appears in a command line or a process listing.
    .PARAMETER IncludePrerelease
        Treat prereleases as valid targets.
    .EXAMPLE
        Get-SpoolerCleanupUpdate
    .EXAMPLE
        if ((Get-SpoolerCleanupUpdate).UpdateAvailable) { Update-SpoolerCleanup -Force }
    .OUTPUTS
        PSCustomObject { CurrentVersion, LatestVersion, UpdateAvailable, Source,
        TagName, ReleaseName, PublishedUtc, ReleaseNotes, DownloadUrl,
        ExpectedSha256, Errors }
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter()]
        [ValidateSet('latest', 'stable', 'prerelease')]
        [string]$Channel = 'latest',

        [Parameter()]
        [AllowNull()][AllowEmptyString()][string]$GitHubToken,

        [Parameter()]
        [switch]$IncludePrerelease
    )

    $null = Get-RuntimeState
    if ([string]::IsNullOrWhiteSpace($GitHubToken)) { $GitHubToken = $env:PSQ_GITHUB_TOKEN }

    Write-LogInfo "Checking GitHub for a newer version (channel: $Channel)..." -Component 'Update'
    $info = Get-UpdateInfoCore -Channel $Channel -Token $GitHubToken -IncludePrerelease:$IncludePrerelease

    if ($info.UpdateAvailable) {
        Write-LogSuccess ("Update available: v{0} -> v{1} (source: {2})" -f $info.CurrentVersion, $info.LatestVersion, $info.Source) -Component 'Update'
    }
    elseif ($info.LatestVersion) {
        Write-LogInfo ("Already current (v{0}); the latest published version is v{1} via {2}." -f $info.CurrentVersion, $info.LatestVersion, $info.Source) -Component 'Update'
    }
    else {
        Write-LogWarn ("The update check was inconclusive: {0}" -f (@($info.Errors) -join ' | ')) -Component 'Update'
    }
    if ($info.ReleaseNotes) { Write-LogDebug ("Release notes:`n{0}" -f $info.ReleaseNotes) -Component 'Update' }
    return $info
}

function Update-SpoolerCleanup {
    <#
    .SYNOPSIS
        Downloads, verifies and installs a newer version of this tool.
    .DESCRIPTION
        The download is verified against the SHA-256 published in the release's
        checksums.txt; on mismatch nothing is installed. New files are STAGED in
        <install>\.update\ and moved into place at the start of the next run,
        because a process must never overwrite the file it is executing from. The
        previous version is kept in <install>\backup-<version>-<stamp>\.
    .PARAMETER Force
        Install even when the discovered version is not newer (re-install / repair).
    .PARAMETER ApplyNow
        Move the staged files into place immediately instead of on next start.
    .PARAMETER SkipChecksumVerification
        Install even when the release publishes no checksums. Deliberately
        explicit: the default is to refuse unverified bytes.
    .PARAMETER InstallDirectory
        Where the tool lives. Defaults to the running script's directory.
    .EXAMPLE
        Update-SpoolerCleanup -Force
    .EXAMPLE
        Update-SpoolerCleanup -Force -ApplyNow
    .OUTPUTS
        PSCustomObject { Updated, Staged, FromVersion, ToVersion, Files,
        BackupPath, VerifyResult, Method, Message }
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    [OutputType([pscustomobject])]
    param(
        [Parameter()]
        [ValidateSet('latest', 'stable', 'prerelease')]
        [string]$Channel = 'latest',

        [Parameter()]
        [switch]$Force,

        [Parameter()]
        [switch]$ApplyNow,

        [Parameter()]
        [switch]$SkipChecksumVerification,

        [Parameter()]
        [AllowNull()][AllowEmptyString()][string]$InstallDirectory,

        [Parameter()]
        [AllowNull()][AllowEmptyString()][string]$GitHubToken,

        [Parameter()]
        [switch]$IncludePrerelease
    )

    $null = Get-RuntimeState
    if ([string]::IsNullOrWhiteSpace($InstallDirectory)) { $InstallDirectory = $Script:RuntimeState.InstallDirectory }
    if ([string]::IsNullOrWhiteSpace($GitHubToken)) { $GitHubToken = $env:PSQ_GITHUB_TOKEN }

    $info = Get-UpdateInfoCore -Channel $Channel -Token $GitHubToken -IncludePrerelease:$IncludePrerelease
    if (-not $info.LatestVersion) {
        throw ("Could not determine the latest version: {0}" -f (@($info.Errors) -join ' | '))
    }

    if (-not $info.UpdateAvailable -and -not $Force) {
        Write-LogInfo "Already up to date (v$($Script:ScriptVersion)); nothing to do. Use -Force to re-install." -Component 'Update'
        return [pscustomobject]@{
            Updated = $false; Staged = $false; FromVersion = $Script:ScriptVersion; ToVersion = $info.LatestVersion
            DownloadedFiles = @(); BackupPath = $null; VerifyResult = 'NotRequired'; Method = $null
            Message = 'Already up to date.'; Files = @()
        }
    }

    $target = "Print Spooler Queue Cleanup in '$InstallDirectory'"
    $action = "Update from v$($Script:ScriptVersion) to v$($info.LatestVersion)"
    if (-not $PSCmdlet.ShouldProcess($target, $action)) {
        Write-LogInfo 'The update was not confirmed; nothing was changed.' -Component 'Update'
        return [pscustomobject]@{
            Updated = $false; Staged = $false; FromVersion = $Script:ScriptVersion; ToVersion = $info.LatestVersion
            DownloadedFiles = @(); BackupPath = $null; VerifyResult = 'NotPerformed'; Method = $null
            Message = 'Cancelled by the operator.'; Files = @()
        }
    }

    $result = Update-SpoolerCleanupCore -UpdateInfo $info -InstallDirectory $InstallDirectory -Force:$Force `
        -SkipChecksumVerification:$SkipChecksumVerification -ApplyNow:$ApplyNow -Token $GitHubToken
    if ($result.Updated -or $result.Staged) { Write-LogSuccess $result.Message -Component 'Update' }
    return $result
}

function Restore-SpoolerBackup {
    <#
    .SYNOPSIS
        Restores a backup archive created by -Backup.
    .DESCRIPTION
        Validates the archive, optionally verifies its SHA-256 against the value
        recorded in the cleanup report, and extracts the spool files plus the JSON
        manifest to a destination directory.

        IMPORTANT: copying spool files back into PRINTERS does NOT re-queue the job.
        The Spooler builds its job table in memory and only reads files it created
        during the current boot. A restore is a forensic recovery of the document
        payload so it can be re-submitted from the application.
    .PARAMETER Path
        Backup ZIP created by -Backup.
    .PARAMETER DestinationDirectory
        Where to extract. Default: "restored_<computer>_<stamp>" next to the archive.
    .PARAMETER ExpectedHash
        SHA-256 the archive must match (BackupHash in the run report).
    .PARAMETER Overwrite
        Overwrite files that already exist in the destination.
    .EXAMPLE
        Restore-SpoolerBackup -Path .\spool-backup_SRV01_20261010-031500.zip
    .EXAMPLE
        Restore-SpoolerBackup -Path $report.Results[0].BackupPath -ExpectedHash $report.Results[0].BackupHash -DestinationDirectory C:\Temp\restored
    .OUTPUTS
        PSCustomObject { BackupPath, Valid, ComputerName, CreatedUtc, ToolVersion,
        ManifestFileCount, FilesRestored, Destination, Sha256, Jobs, Note }
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [Alias('BackupPath', 'FullName')]
        [AllowEmptyString()][string]$Path,

        [Parameter()]
        [Alias('Destination')]
        [AllowNull()][AllowEmptyString()][string]$DestinationDirectory,

        [Parameter()]
        [AllowNull()][AllowEmptyString()][string]$ExpectedHash,

        [Parameter()]
        [switch]$Overwrite,

        [Parameter()]
        [switch]$Force
    )

    process {
        $null = Get-RuntimeState
        $info = Get-SpoolBackupInfo -Path $Path
        if (-not $info.Valid) { throw "Backup '$Path' is not readable: $($info.Error)" }

        $dest = if ($DestinationDirectory) { $DestinationDirectory } else { "<next to the archive>\restored_$($info.ComputerName)_<stamp>" }
        $action = "Restore $($info.FileCount) spool file(s) from '$($info.ComputerName)' to '$dest'"
        if (-not ($Force -or $PSCmdlet.ShouldProcess($Path, $action))) {
            Write-LogInfo 'The restore was not confirmed; nothing was extracted.' -Component 'Restore'
            return $null
        }

        $params = @{ Path = $Path; Overwrite = $Overwrite; Force = $Force }
        if (-not [string]::IsNullOrWhiteSpace($ExpectedHash)) { $params['ExpectedHash'] = $ExpectedHash }
        if (-not [string]::IsNullOrWhiteSpace($DestinationDirectory)) { $params['DestinationDirectory'] = $DestinationDirectory }

        $result = Restore-SpoolerBackupInternal @params
        Write-LogSuccess ("Restored {0} file(s) to '{1}'." -f $result.FilesRestored, $result.Destination) -Component 'Restore'
        return $result
    }
}

function Get-SpoolerBackup {
    <#
    .SYNOPSIS
        Lists the backup archives in a backup directory.
    .PARAMETER BackupPath
        Directory to scan. Default: %ProgramData%\PrintSpoolerCleanup\Backups.
    .EXAMPLE
        Get-SpoolerBackup | Format-Table CreatedUtc, ComputerName, FileCount, SizeBytes
    .OUTPUTS
        PSCustomObject[]
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject[]])]
    param(
        [Parameter(Position = 0)]
        [AllowNull()][AllowEmptyString()][string]$BackupPath
    )
    if ([string]::IsNullOrWhiteSpace($BackupPath)) {
        $root = if ($env:ProgramData) { $env:ProgramData } else { [IO.Path]::GetTempPath() }
        $BackupPath = Join-Path (Join-Path $root $Script:ProductFamily) 'Backups'
    }
    return @(Get-SpoolBackupInventory -BackupPath $BackupPath)
}

function Install-SpoolerCleanupTask {
    <#
    .SYNOPSIS
        Registers a Task Scheduler task that maintains the print queue automatically.
    .DESCRIPTION
        Always launches the non-interactive PowerShell entry point with
        -Force -NonInteractive -Quiet -NoProgress -Backup -AutoUpdate, so the task
        can never block on a prompt - the failure mode of the schtasks one-liner
        that earlier documentation recommended.

        Supports Daily, Weekly, Hourly, AtStartup, OnEvent and Once schedules, an
        execution time limit, restart-on-failure and "start when available", and
        disables the battery restrictions that silently skip maintenance on laptops.
    .PARAMETER ScriptPath
        Full path to Clear-PrintSpoolerQueue.ps1. Required.
    .PARAMETER TaskName
        Task name; may include a folder ("PrintSpoolerCleanup\Nightly").
    .PARAMETER Schedule
        Daily (default), Weekly, Hourly, AtStartup, OnEvent, Once.
    .PARAMETER StartTime
        Local start time as HH:mm (default 03:00).
    .PARAMETER IntervalHours
        Interval for the Hourly schedule.
    .PARAMETER ComputerName
        Targets the task should clean. Omit for the local machine only.
    .PARAMETER ComputerListFile
        Text file with one computer per line. Better for large sites: the list can
        be updated without re-registering the task.
    .PARAMETER Arguments
        Extra arguments appended verbatim, e.g. '-OlderThan 6h -Repair'.
    .PARAMETER RunAsAccount
        Default: NT AUTHORITY\SYSTEM.
    .PARAMETER EventLogName
        For -Schedule OnEvent: the log to subscribe to.
    .PARAMETER EventSource
        For -Schedule OnEvent: the provider name.
    .PARAMETER EventId
        For -Schedule OnEvent: the event id. 7031 from "Service Control Manager"
        in the System log is the classic "the spooler crashed" trigger.
    .PARAMETER ExecutionTimeLimitMinutes
        Default 30. The task is terminated if it exceeds this.
    .PARAMETER Method
        Auto (default), Module or Xml.
    .EXAMPLE
        Install-SpoolerCleanupTask -ScriptPath C:\Tools\Clear-PrintSpoolerQueue.ps1 -Schedule Daily -StartTime 03:00 -PassThru
    .EXAMPLE
        Install-SpoolerCleanupTask -ScriptPath C:\Tools\Clear-PrintSpoolerQueue.ps1 -Schedule OnEvent -EventLogName System -EventSource "Service Control Manager" -EventId 7031
    .EXAMPLE
        Install-SpoolerCleanupTask -ScriptPath C:\Tools\Clear-PrintSpoolerQueue.ps1 -Schedule Hourly -IntervalHours 6 -ComputerListFile C:\Tools\hosts.txt -Arguments '-OlderThan 2h'
    .OUTPUTS
        Boolean, or the task definition with -PassThru.
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    [OutputType([bool], [pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()][string]$ScriptPath,

        [Parameter()]
        [AllowEmptyString()][string]$TaskName = 'PrintSpoolerCleanup\Clear Stuck Print Queue',

        [Parameter()]
        [ValidateSet('Daily', 'Weekly', 'Hourly', 'AtStartup', 'OnEvent', 'Once')]
        [string]$Schedule = 'Daily',

        [Parameter()]
        [ValidatePattern('^([01]\d|2[0-3]):[0-5]\d$')]
        [string]$StartTime = '03:00',

        [Parameter()]
        [ValidateRange(1, 168)]
        [int]$IntervalHours = 6,

        [Parameter()]
        [AllowNull()][AllowEmptyCollection()][string[]]$ComputerName,

        [Parameter()]
        [AllowNull()][AllowEmptyString()][string]$ComputerListFile,

        [Parameter()]
        [AllowNull()][AllowEmptyString()][string]$Arguments,

        [Parameter()]
        [AllowEmptyString()][string]$RunAsAccount = 'NT AUTHORITY\SYSTEM',

        [Parameter()]
        [AllowNull()][AllowEmptyString()][string]$EventLogName,

        [Parameter()]
        [AllowNull()][AllowEmptyString()][string]$EventSource,

        [Parameter()]
        [AllowNull()][Nullable[int]]$EventId,

        [Parameter()]
        [ValidateRange(5, 1440)]
        [int]$ExecutionTimeLimitMinutes = 30,

        [Parameter()]
        [switch]$RunOnlyIfIdle,

        [Parameter()]
        [switch]$NoBackup,

        [Parameter()]
        [switch]$NoAutoUpdate,

        [Parameter()]
        [ValidateSet('Auto', 'Module', 'Xml')]
        [string]$Method = 'Auto',

        [Parameter()]
        [switch]$PassThru
    )

    $action = "Register a $Schedule scheduled task running '$ScriptPath' as $RunAsAccount"
    if (-not $PSCmdlet.ShouldProcess($TaskName, $action)) { return $false }

    $params = @{
        TaskName                  = $TaskName
        ScriptPath                = $ScriptPath
        Schedule                  = $Schedule
        StartTime                 = $StartTime
        IntervalHours             = $IntervalHours
        RunAsAccount              = $RunAsAccount
        ExecutionTimeLimitMinutes = $ExecutionTimeLimitMinutes
        Method                    = $Method
        Backup                    = (-not $NoBackup)
        AutoUpdate                = (-not $NoAutoUpdate)
    }
    if ($PassThru) { $params['PassThru'] = $true }
    if ($RunOnlyIfIdle) { $params['RunOnlyIfIdle'] = $true }
    if ($ComputerName -and @($ComputerName).Count -gt 0) { $params['ComputerName'] = $ComputerName }
    if (-not [string]::IsNullOrWhiteSpace($ComputerListFile)) { $params['ComputerListFile'] = $ComputerListFile }
    if (-not [string]::IsNullOrWhiteSpace($Arguments)) { $params['Arguments'] = $Arguments }
    if (-not [string]::IsNullOrWhiteSpace($EventLogName)) { $params['EventLogName'] = $EventLogName }
    if (-not [string]::IsNullOrWhiteSpace($EventSource)) { $params['EventSource'] = $EventSource }
    if ($null -ne $EventId) { $params['EventId'] = $EventId }

    return Install-SpoolerCleanupTaskInternal @params
}

function Uninstall-SpoolerCleanupTask {
    <#
    .SYNOPSIS
        Removes the scheduled task registered by Install-SpoolerCleanupTask.
    .EXAMPLE
        Uninstall-SpoolerCleanupTask
        Uninstall-SpoolerCleanupTask -TaskName 'PrintSpoolerCleanup\Hourly'
    .OUTPUTS
        System.Boolean
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    [OutputType([bool])]
    param(
        [Parameter(Position = 0)]
        [AllowEmptyString()][string]$TaskName = 'PrintSpoolerCleanup\Clear Stuck Print Queue'
    )
    if (-not $PSCmdlet.ShouldProcess($TaskName, 'Delete the scheduled task')) { return $false }

    $ok = Uninstall-SpoolerCleanupTaskInternal -TaskName $TaskName
    if ($ok) { Write-LogSuccess "Scheduled task '$TaskName' removed." -Component 'Task' }
    else { Write-LogWarn "Scheduled task '$TaskName' could not be removed (it may not exist)." -Component 'Task' }
    return $ok
}

function Get-SpoolerCleanupTask {
    <#
    .SYNOPSIS
        Shows the registered task, its last result and its next run time.
    .EXAMPLE
        Get-SpoolerCleanupTask
        Get-SpoolerCleanupTask | Format-List *
    .OUTPUTS
        PSCustomObject
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Position = 0)]
        [AllowEmptyString()][string]$TaskName = 'PrintSpoolerCleanup\Clear Stuck Print Queue'
    )
    $task = Get-SpoolerCleanupTaskInternal -TaskName $TaskName
    if ($task.Exists -and $task.PSObject.Properties['LastTaskResult']) {
        $meaning = Get-ExitCodeMeaning -Code ([int]$task.LastTaskResult)
        $text = if ($meaning) { $meaning.Name } else { "Unknown (0x{0:X})" -f [int]$task.LastTaskResult }
        $task | Add-Member -NotePropertyName LastTaskResultMeaning -NotePropertyValue $text -Force
    }
    return $task
}

function Start-SpoolerCleanupTask {
    <#
    .SYNOPSIS
        Runs the registered task immediately and optionally waits for the result.
    .PARAMETER Wait
        Block until the task finishes and return its exit code.
    .PARAMETER TimeoutSeconds
        Maximum time to wait with -Wait. Default 600.
    .EXAMPLE
        Start-SpoolerCleanupTask -Wait
    .OUTPUTS
        PSCustomObject { Started, TaskName, ExitCode, Waited }
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Position = 0)]
        [AllowEmptyString()][string]$TaskName = 'PrintSpoolerCleanup\Clear Stuck Print Queue',

        [Parameter()]
        [switch]$Wait,

        [Parameter()]
        [ValidateRange(10, 7200)]
        [int]$TimeoutSeconds = 600
    )
    return Start-SpoolerCleanupTaskInternal -TaskName $TaskName -Wait:$Wait -TimeoutSeconds $TimeoutSeconds
}

function Get-ExitCodeMeaning {
    <#
    .SYNOPSIS
        Translates a numeric exit code into its contract name and description.
    .EXAMPLE
        Get-ExitCodeMeaning -Code 12
    .EXAMPLE
        Get-ExitCodeMeaning
        Lists the whole contract.
    .OUTPUTS
        PSCustomObject { Code, Name, Description }
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Position = 0)][AllowNull()][Nullable[int]]$Code)

    $descriptions = @{
        Success             = 'All targets were processed and the Spooler verified running.'
        PartialSuccess      = 'Some targets succeeded, others failed or produced warnings.'
        Failed              = 'No target could be processed.'
        BadArguments        = 'Invalid parameter combination or unreadable input.'
        Cancelled           = 'The operator declined the confirmation prompt.'
        NotAdministrator    = 'Elevation is required and was not available.'
        UpdateCheckFailed   = 'GitHub could not be reached during -CheckUpdate.'
        UpdateAvailable     = '-CheckUpdate found a newer version. Nothing was changed.'
        UpdateApplied       = 'A newer version was installed.'
        UpdateFailed        = 'Update download, verification or staging failed.'
        UnsupportedPlatform = 'Not Windows, or the Print Spooler feature is not installed.'
        DiagnosticsOnly     = '-Diagnostic completed. Informational, not an error.'
        ServiceFailure      = 'The Spooler could not be stopped or restarted.'
        RestoreFailed       = 'Restore-SpoolerBackup could not complete.'
        SignatureInvalid    = '-VerifySignature found a broken, missing or unexpected signature.'
        Timeout             = 'The watchdog expired and the run was terminated.'
        Aborted             = 'Interrupted with Ctrl+C or by host shutdown.'
    }

    $rows = @(foreach ($name in @($Script:ExitCode.Keys)) {
            [pscustomobject]@{
                Code        = [int]$Script:ExitCode[$name]
                Name        = $name
                Description = if ($descriptions.Contains($name)) { $descriptions[$name] } else { '' }
            }
        })

    if ($null -eq $Code) { return ($rows | Sort-Object Code) }
    return @($rows | Where-Object { $_.Code -eq $Code } | Select-Object -First 1)
}

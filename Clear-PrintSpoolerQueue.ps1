<#
.SYNOPSIS
    Forcibly clears the Windows Print Spooler queue on local or remote computers
    with logging, error handling, and safe service stop/start orchestration.

.DESCRIPTION
    Senior-level system-administration utility for clearing stuck print jobs.
    Compatible with Windows PowerShell 5.1 and PowerShell 7+.

    Behavior:
    - Stops the Spooler service safely (with timeout and recovery on failure).
    - Removes residual *.SHD / *.SPL files from the active spool directory.
    - Restarts the Spooler service and verifies its running state.
    - Supports local execution and remote execution via Invoke-Command (WinRM).
    - Honors -WhatIf / -Confirm semantics on the destructive cleanup step.

.PARAMETER ComputerName
    One or more target computers. Defaults to the local machine.
    Pipeline input is supported.

.PARAMETER LogPath
    Full path to the log file. If omitted, a timestamped log file is created
    next to the script (Clear-PrintSpoolerQueue_yyyyMMdd-HHmmss.log).

.PARAMETER Force
    Skip the interactive confirmation prompt.

.PARAMETER ServiceTimeoutSeconds
    Maximum time, in seconds, to wait for the Spooler service to reach the
    desired state (Stopped or Running). Default: 30.

.EXAMPLE
    .\Clear-PrintSpoolerQueue.ps1
    Cleans the local spooler queue with confirmation.

.EXAMPLE
    .\Clear-PrintSpoolerQueue.ps1 -ComputerName "SERVER01","PC-FINANCE" -Force
    Bulk cleanup on two remote machines without prompting (requires WinRM).

.EXAMPLE
    .\Clear-PrintSpoolerQueue.ps1 -WhatIf
    Reports what would happen without changing anything.

.EXAMPLE
    .\Clear-PrintSpoolerQueue.ps1 -LogPath 'C:\Admin\Logs\SpoolerCleanup.log' -Verbose
    Runs with verbose tracing and a custom log path.

.NOTES
    Author : Mikhail Deynekin
    Email  : Mikhail@Deynekin.com
    Site   : https://Deynekin.com
    GitHub : https://github.com/paulmann/Print-Spooler-Queue-Cleanup

    Script Version: 3.1.0 (full rewrite — fixes invalid top-level begin/process/end
    blocks, removes PS7-incompatible -ComputerName usage on Get-Service/Stop-Service,
    fixes Clear-SpoolFiles remote scriptblock variable scoping, hardens error handling).

    Requirements:
    - Windows PowerShell 5.1 or PowerShell 7+
    - Administrator privileges for local execution
    - WinRM enabled on remote targets (Invoke-Command transport)
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
    [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$ComputerName = @($env:COMPUTERNAME),

    [Parameter()]
    [string]$LogPath,

    [Parameter()]
    [switch]$Force,

    [Parameter()]
    [ValidateRange(5, 600)]
    [int]$ServiceTimeoutSeconds = 30
)

begin {
    # NOTE: In a PowerShell *script*, the named blocks begin/process/end must
    # follow the param() block immediately, with no other statements between
    # them. Set-StrictMode and any other initialization therefore live INSIDE
    # the begin block, not between param and begin.
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    $script:ScriptVersion = '3.1.0'
    $script:StartTime     = Get-Date
    $script:LocalNames    = @('localhost', '.', '127.0.0.1', $env:COMPUTERNAME) |
                            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    $script:Config = @{
        ServiceName       = 'Spooler'
        RequiredPSVersion = [Version]'5.1'
        SettleSeconds     = 3
        LogPrefix         = '[Print Spooler Cleanup]'
    }

    $script:Stats = [ordered]@{
        StartTime         = $script:StartTime
        EndTime           = $null
        ComputersAttempted = 0
        ComputersSucceeded = 0
        ComputersFailed    = 0
        FilesProcessed     = 0
        JobsCleared        = 0
        ErrorsEncountered  = 0
        OperationResult    = 'Unknown'
    }

    # Resolve script directory robustly (works when dot-sourced, run via -File, etc.)
    $invocationPath = $MyInvocation.MyCommand.Path
    if ($invocationPath) {
        $script:ScriptDirectory = Split-Path -Parent $invocationPath
    } elseif ($PSScriptRoot) {
        $script:ScriptDirectory = $PSScriptRoot
    } else {
        $script:ScriptDirectory = (Get-Location).Path
    }

    if (-not $LogPath) {
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $LogPath = Join-Path -Path $script:ScriptDirectory -ChildPath "Clear-PrintSpoolerQueue_$stamp.log"
    }
    $script:LogPath = $LogPath

    #region Helper functions

    # Function Version: 1.0.0
    function Write-Log {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)][string]$Message,
            [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')][string]$Level = 'INFO'
        )
        $entry = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
        Write-Verbose $entry
        try {
            Add-Content -Path $script:LogPath -Value $entry -Encoding UTF8 -ErrorAction Stop
        } catch {
            # Never let logging failures abort the run; surface once on the host.
            Write-Warning ("Log write failed ({0}): {1}" -f $script:LogPath, $_.Exception.Message)
        }
    }

    # Function Version: 1.0.0
    function Write-StatusMessage {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)][string]$Message,
            [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Progress')][string]$Level = 'Info'
        )
        $iconMap  = @{ Info = 'i'; Success = 'OK'; Warning = '!'; Error = 'X'; Progress = '>' }
        $colorMap = @{ Info = 'White'; Success = 'Green'; Warning = 'Yellow'; Error = 'Red'; Progress = 'Cyan' }
        $line = "{0} [{1}] {2} {3}" -f $script:Config.LogPrefix, (Get-Date -Format 'HH:mm:ss'), $iconMap[$Level], $Message
        Write-Host $line -ForegroundColor $colorMap[$Level]
        $logLevel = switch ($Level) { 'Warning' { 'WARN' } 'Error' { 'ERROR' } default { 'INFO' } }
        Write-Log -Message $Message -Level $logLevel
    }

    # Function Version: 1.0.0
    function Write-HeaderMessage {
        $bar = ('=' * 78)
        $lines = @(
            $bar,
            "    Print Spooler Queue Cleanup Utility v$script:ScriptVersion",
            "    Author : Mikhail Deynekin <Mikhail@Deynekin.com>",
            "    Site   : https://Deynekin.com",
            "    GitHub : https://github.com/paulmann/Print-Spooler-Queue-Cleanup",
            $bar,
            ''
        )
        $lines | ForEach-Object { Write-Host $_ -ForegroundColor Cyan }
    }

    # Function Version: 1.0.0
    function Test-AdministratorRights {
        try {
            if ($IsLinux -or $IsMacOS) { return $false }
            $principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
            return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        } catch {
            Write-StatusMessage "Failed to check administrator rights: $($_.Exception.Message)" -Level Error
            return $false
        }
    }

    # Function Version: 1.0.0
    function Test-PowerShellVersion {
        $current = $PSVersionTable.PSVersion
        if ($current -lt $script:Config.RequiredPSVersion) {
            Write-StatusMessage "PowerShell $current is below required $($script:Config.RequiredPSVersion)" -Level Error
            return $false
        }
        return $true
    }

    # Function Version: 1.0.0
    function Test-IsLocalTarget {
        param([Parameter(Mandatory)][string]$Computer)
        return ($script:LocalNames -contains $Computer)
    }

    # Function Version: 1.1.0
    # Cross-version safe: prefer Invoke-Command (works on PS 5.1 and 7+).
    # Get-Service/Stop-Service -ComputerName were removed in PowerShell 7,
    # so we route all remote operations through Invoke-Command.
    function Invoke-OnTarget {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)][string]$Computer,
            [Parameter(Mandatory)][scriptblock]$ScriptBlock,
            [object[]]$ArgumentList = @()
        )
        if (Test-IsLocalTarget -Computer $Computer) {
            if ($ArgumentList.Count -gt 0) {
                return & $ScriptBlock @ArgumentList
            }
            return & $ScriptBlock
        }
        $params = @{
            ComputerName = $Computer
            ScriptBlock  = $ScriptBlock
            ErrorAction  = 'Stop'
        }
        if ($ArgumentList.Count -gt 0) { $params['ArgumentList'] = $ArgumentList }
        return Invoke-Command @params
    }

    # Function Version: 1.0.0
    function Get-SpoolerDirectoryPath {
        param([Parameter(Mandatory)][string]$Computer)
        $sb = {
            $regKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers'
            $custom = (Get-ItemProperty -Path $regKey -Name 'DefaultSpoolDirectory' -ErrorAction SilentlyContinue).DefaultSpoolDirectory
            if ([string]::IsNullOrWhiteSpace($custom)) { "$env:SystemRoot\System32\spool\PRINTERS" } else { $custom }
        }
        return Invoke-OnTarget -Computer $Computer -ScriptBlock $sb
    }

    # Function Version: 1.0.0
    function Get-SpoolFileStatistics {
        param(
            [Parameter(Mandatory)][string]$Computer,
            [Parameter(Mandatory)][string]$SpoolDir
        )
        $sb = {
            param($Path)
            if (-not (Test-Path -LiteralPath $Path)) {
                return [pscustomobject]@{ Files = 0; SHDFiles = 0; SPLFiles = 0; TotalSize = 0L; Path = $Path; Exists = $false }
            }
            $all = @(Get-ChildItem -LiteralPath $Path -File -Force -ErrorAction SilentlyContinue)
            $shd = @($all | Where-Object { $_.Extension -ieq '.SHD' }).Count
            $spl = @($all | Where-Object { $_.Extension -ieq '.SPL' }).Count
            $sum = ($all | Measure-Object -Property Length -Sum).Sum
            if (-not $sum) { $sum = 0L }
            [pscustomobject]@{ Files = $all.Count; SHDFiles = $shd; SPLFiles = $spl; TotalSize = [int64]$sum; Path = $Path; Exists = $true }
        }
        return Invoke-OnTarget -Computer $Computer -ScriptBlock $sb -ArgumentList @($SpoolDir)
    }

    # Function Version: 1.1.0
    # Stops Spooler via Invoke-Command on remote, native cmdlets locally.
    function Stop-SpoolerServiceOnTarget {
        param(
            [Parameter(Mandatory)][string]$Computer,
            [int]$TimeoutSeconds = 30
        )
        Write-StatusMessage "Stopping Print Spooler service on '$Computer'..." -Level Progress
        $serviceName = $script:Config.ServiceName
        $sb = {
            param($name, $timeout)
            $svc = Get-Service -Name $name -ErrorAction Stop
            if ($svc.Status -eq 'Stopped') { return @{ Result = 'AlreadyStopped' } }
            Stop-Service -Name $name -Force -ErrorAction Stop
            try {
                $svc.WaitForStatus('Stopped', [TimeSpan]::FromSeconds($timeout))
            } catch {
                return @{ Result = 'Timeout'; Error = $_.Exception.Message }
            }
            return @{ Result = 'Stopped' }
        }
        try {
            $r = Invoke-OnTarget -Computer $Computer -ScriptBlock $sb -ArgumentList @($serviceName, $TimeoutSeconds)
            switch ($r.Result) {
                'AlreadyStopped' { Write-StatusMessage "Print Spooler was already stopped on '$Computer'" -Level Info; return $true }
                'Stopped'        { Write-StatusMessage "Print Spooler stopped successfully on '$Computer'"  -Level Success; return $true }
                'Timeout'        { throw "Service did not reach 'Stopped' within ${TimeoutSeconds}s on '$Computer': $($r.Error)" }
                default          { throw "Unexpected stop result on '$Computer': $($r.Result)" }
            }
        } catch {
            Write-StatusMessage "Failed to stop Print Spooler on '$Computer': $($_.Exception.Message)" -Level Error
            $script:Stats.ErrorsEncountered++
            return $false
        }
    }

    # Function Version: 1.1.0
    # Targets only *.SHD / *.SPL spool files (non-destructive: leaves any
    # other content in the directory untouched). Honors -WhatIf via the caller.
    function Clear-SpoolFilesOnTarget {
        param(
            [Parameter(Mandatory)][string]$Computer,
            [Parameter(Mandatory)][string]$SpoolDir
        )
        Write-StatusMessage "Analyzing spool directory '$SpoolDir' on '$Computer'..." -Level Progress
        $pre = Get-SpoolFileStatistics -Computer $Computer -SpoolDir $SpoolDir
        if (-not $pre.Exists) {
            Write-StatusMessage "Spool directory '$SpoolDir' not found on '$Computer'" -Level Warning
            return $false
        }
        if ($pre.Files -eq 0) {
            Write-StatusMessage "Spool directory is already empty on '$Computer'" -Level Info
            return $true
        }
        Write-StatusMessage ("Found {0} files ({1} .SHD, {2} .SPL, {3} bytes) on '{4}'" -f $pre.Files, $pre.SHDFiles, $pre.SPLFiles, $pre.TotalSize, $Computer) -Level Info

        $sb = {
            param($Path)
            $items = @(Get-ChildItem -LiteralPath $Path -File -Force -ErrorAction SilentlyContinue |
                       Where-Object { $_.Extension -ieq '.SHD' -or $_.Extension -ieq '.SPL' })
            $out = New-Object System.Collections.Generic.List[object]
            foreach ($f in $items) {
                try {
                    Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop
                    $out.Add([pscustomobject]@{ Success = $true; File = $f.Name; Error = $null }) | Out-Null
                } catch {
                    $out.Add([pscustomobject]@{ Success = $false; File = $f.Name; Error = $_.Exception.Message }) | Out-Null
                }
            }
            return ,$out.ToArray()
        }

        Write-StatusMessage "Clearing spool files on '$Computer'..." -Level Progress
        $results = @(Invoke-OnTarget -Computer $Computer -ScriptBlock $sb -ArgumentList @($SpoolDir))
        $ok   = @($results | Where-Object { $_.Success }).Count
        $bad  = @($results | Where-Object { -not $_.Success })
        if ($bad.Count -gt 0) {
            foreach ($b in $bad) {
                Write-StatusMessage "Failed to remove '$($b.File)': $($b.Error)" -Level Warning
            }
            $script:Stats.ErrorsEncountered += $bad.Count
        }
        $script:Stats.FilesProcessed += $ok
        $script:Stats.JobsCleared    += [Math]::Max($pre.SHDFiles, $pre.SPLFiles)
        Write-StatusMessage ("Removed {0} of {1} spool files on '{2}'" -f $ok, $results.Count, $Computer) -Level Success
        return ($bad.Count -eq 0)
    }

    # Function Version: 1.1.0
    function Start-SpoolerServiceOnTarget {
        param(
            [Parameter(Mandatory)][string]$Computer,
            [int]$TimeoutSeconds = 30
        )
        Write-StatusMessage "Starting Print Spooler service on '$Computer'..." -Level Progress
        $serviceName = $script:Config.ServiceName
        $sb = {
            param($name, $timeout)
            $svc = Get-Service -Name $name -ErrorAction Stop
            if ($svc.Status -ne 'Running') {
                Start-Service -Name $name -ErrorAction Stop
            }
            try {
                $svc.WaitForStatus('Running', [TimeSpan]::FromSeconds($timeout))
            } catch {
                return @{ Result = 'Timeout'; Error = $_.Exception.Message }
            }
            return @{ Result = 'Running' }
        }
        try {
            $r = Invoke-OnTarget -Computer $Computer -ScriptBlock $sb -ArgumentList @($serviceName, $TimeoutSeconds)
            switch ($r.Result) {
                'Running' { Write-StatusMessage "Print Spooler is running on '$Computer'" -Level Success; return $true }
                'Timeout' { throw "Service did not reach 'Running' within ${TimeoutSeconds}s on '$Computer': $($r.Error)" }
                default   { throw "Unexpected start result on '$Computer': $($r.Result)" }
            }
        } catch {
            Write-StatusMessage "Failed to start Print Spooler on '$Computer': $($_.Exception.Message)" -Level Error
            $script:Stats.ErrorsEncountered++
            return $false
        }
    }

    # Function Version: 1.0.0
    function Show-ExecutionStatistics {
        $script:Stats.EndTime = Get-Date
        $duration = $script:Stats.EndTime - $script:Stats.StartTime
        $bar = ('=' * 78)
        Write-Host ''
        Write-Host $bar -ForegroundColor Cyan
        Write-Host '    EXECUTION SUMMARY' -ForegroundColor Cyan
        Write-Host $bar -ForegroundColor Cyan
        $rows = @(
            @{ L = 'Execution Time';     V = ('{0:hh\:mm\:ss\.fff}' -f $duration); C = 'White' },
            @{ L = 'Computers Attempted'; V = $script:Stats.ComputersAttempted; C = 'White' },
            @{ L = 'Computers Succeeded'; V = $script:Stats.ComputersSucceeded; C = if ($script:Stats.ComputersSucceeded -gt 0) { 'Green' } else { 'Yellow' } },
            @{ L = 'Computers Failed';    V = $script:Stats.ComputersFailed;    C = if ($script:Stats.ComputersFailed -eq 0) { 'Green' } else { 'Red' } },
            @{ L = 'Print Jobs Cleared'; V = $script:Stats.JobsCleared;        C = if ($script:Stats.JobsCleared -gt 0) { 'Green' } else { 'Yellow' } },
            @{ L = 'Files Processed';    V = $script:Stats.FilesProcessed;     C = if ($script:Stats.FilesProcessed -gt 0) { 'Green' } else { 'Yellow' } },
            @{ L = 'Errors Encountered'; V = $script:Stats.ErrorsEncountered;  C = if ($script:Stats.ErrorsEncountered -eq 0) { 'Green' } else { 'Red' } },
            @{ L = 'Operation Result';   V = $script:Stats.OperationResult;    C = if ($script:Stats.OperationResult -eq 'Success') { 'Green' } else { 'Red' } }
        )
        foreach ($row in $rows) {
            $pad = ' ' * [Math]::Max(1, 22 - $row.L.Length)
            Write-Host ("    {0}:{1}" -f $row.L, $pad) -NoNewline -ForegroundColor Gray
            Write-Host $row.V -ForegroundColor $row.C
        }
        Write-Host $bar -ForegroundColor Cyan
    }

    #endregion Helper functions

    Write-HeaderMessage
    Write-Log -Message "Script v$script:ScriptVersion started on PowerShell $($PSVersionTable.PSVersion)."

    if (-not (Test-PowerShellVersion)) { throw "PowerShell version requirements not met." }

    # When values arrive via the pipeline, $ComputerName here still holds the
    # parameter's default — a real check for "do we touch the local box?" runs
    # again per item in the process block. We skip empty strings to be safe on
    # hosts where $env:COMPUTERNAME is unexpectedly unset.
    $touchesLocal = $false
    foreach ($c in $ComputerName) {
        if ([string]::IsNullOrWhiteSpace($c)) { continue }
        if (Test-IsLocalTarget -Computer $c) { $touchesLocal = $true; break }
    }
    if ($touchesLocal -and -not (Test-AdministratorRights)) {
        throw "Administrator privileges are required to manage the local Print Spooler service."
    }
}

process {
    foreach ($Computer in $ComputerName) {
        if ([string]::IsNullOrWhiteSpace($Computer)) {
            Write-StatusMessage "Skipping empty computer name" -Level Warning
            continue
        }
        $script:Stats.ComputersAttempted++
        Write-StatusMessage "Beginning cleanup on '$Computer'" -Level Info
        $thisOk = $true

        try {
            $spoolDir = Get-SpoolerDirectoryPath -Computer $Computer
            Write-Log -Message "Spool directory on '$Computer': '$spoolDir'"

            $confirmTarget = "Print Spooler queue on '$Computer' (directory: $spoolDir)"
            $confirmAction = "Stop service, delete *.SHD/*.SPL files, restart service"
            if (-not $Force -and -not $PSCmdlet.ShouldProcess($confirmTarget, $confirmAction)) {
                Write-StatusMessage "Skipped '$Computer' (not confirmed)" -Level Info
                $script:Stats.ComputersFailed++
                continue
            }

            if (-not (Stop-SpoolerServiceOnTarget -Computer $Computer -TimeoutSeconds $ServiceTimeoutSeconds)) {
                throw "Could not stop Spooler on '$Computer'."
            }

            Start-Sleep -Seconds $script:Config.SettleSeconds

            if (-not (Clear-SpoolFilesOnTarget -Computer $Computer -SpoolDir $spoolDir)) {
                Write-StatusMessage "Spool file cleanup completed with errors on '$Computer'" -Level Warning
                $thisOk = $false
            }

            if (-not (Start-SpoolerServiceOnTarget -Computer $Computer -TimeoutSeconds $ServiceTimeoutSeconds)) {
                throw "Could not start Spooler on '$Computer'."
            }

            if ($thisOk) {
                $script:Stats.ComputersSucceeded++
                Write-StatusMessage "Cleanup completed successfully on '$Computer'" -Level Success
            } else {
                $script:Stats.ComputersFailed++
                Write-StatusMessage "Cleanup completed with warnings on '$Computer'" -Level Warning
            }
        }
        catch {
            $script:Stats.ComputersFailed++
            $script:Stats.ErrorsEncountered++
            Write-StatusMessage "Critical error on '$Computer': $($_.Exception.Message)" -Level Error

            # Best-effort recovery: try to leave Spooler running so the host
            # is not left without printing capability.
            try {
                Write-StatusMessage "Attempting Spooler recovery on '$Computer'..." -Level Warning
                $null = Start-SpoolerServiceOnTarget -Computer $Computer -TimeoutSeconds $ServiceTimeoutSeconds
            } catch {
                Write-StatusMessage "Spooler recovery failed on '$Computer': $($_.Exception.Message)" -Level Error
            }
        }
    }
}

end {
    if ($script:Stats.ComputersAttempted -gt 0 -and $script:Stats.ComputersFailed -eq 0) {
        $script:Stats.OperationResult = 'Success'
    } elseif ($script:Stats.ComputersSucceeded -gt 0) {
        $script:Stats.OperationResult = 'PartialSuccess'
    } else {
        $script:Stats.OperationResult = 'Failed'
    }

    Show-ExecutionStatistics
    Write-Log -Message ("Run finished. Result={0}; Succeeded={1}; Failed={2}; Errors={3}." -f `
        $script:Stats.OperationResult, $script:Stats.ComputersSucceeded, $script:Stats.ComputersFailed, $script:Stats.ErrorsEncountered)

    # Map result to a meaningful exit code without using `exit` (which would
    # nuke the host when the script is dot-sourced). $LASTEXITCODE-style hosts
    # can read this via $? / the process exit when invoked by powershell.exe -File.
    switch ($script:Stats.OperationResult) {
        'Success'        { $global:LASTEXITCODE = 0 }
        'PartialSuccess' { $global:LASTEXITCODE = 2 }
        default          { $global:LASTEXITCODE = 1 }
    }
}

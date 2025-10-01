<#
.SYNOPSIS
    Advanced PowerShell script for forcibly clearing Windows Print Spooler queue
    with comprehensive logging, error handling, and professional system administration features.

.DESCRIPTION
    This script provides enterprise-grade functionality for clearing print queue stuck jobs.
    Fully compatible with PowerShell 5.1 and PowerShell 7+.
    
    Key Features:
    - Real-time status monitoring with beautiful, color-coded console output (emojis & colors)
    - Comprehensive error handling with detailed logging and service recovery attempts
    - Detailed statistics tracking (jobs cleared, files processed, execution time)
    - Administrator privilege and PowerShell version verification
    - Service state validation with timeout handling
    - Professional header and progress indication for all major operations
    - Modular design with clear separation of concerns
    - Support for remote execution on multiple computers (ComputerName parameter)
    - WhatIf and Confirm support for safe execution planning

.PARAMETER ComputerName
    Specifies the target computer(s) on which to perform the cleanup.
    Accepts an array of computer names or IP addresses for bulk operations.
    The default is the local computer ('localhost').

.PARAMETER LogPath
    Specifies the full path for the log file.
    If not provided, a timestamped log file (e.g., Clear-PrintSpoolerQueue_20240521-120000.log)
    will be created in the same directory as the script.

.PARAMETER Force
    Skips the interactive confirmation prompt before performing the cleanup.
    Use this switch for automated or unattended execution.

.EXAMPLE
    .\Clear-PrintSpoolerQueue.ps1
    Executes the print spooler cleanup on the local machine with confirmation prompt.

.EXAMPLE
    .\Clear-PrintSpoolerQueue.ps1 -ComputerName "PRINT-SRV01", "PRINT-SRV02" -Force
    Clears the print queue on two remote servers without prompting for confirmation.

.EXAMPLE
    .\Clear-PrintSpoolerQueue.ps1 -WhatIf
    Shows what actions would be performed without making any changes to the system.

.EXAMPLE
    .\Clear-PrintSpoolerQueue.ps1 -LogPath "C:\Admin\Logs\SpoolerCleanup.log" -Verbose
    Runs the cleanup with verbose output and saves the log to a custom location.

.AUTHOR
    Mikhail Deynekin (mid1977@gmail.com)
    Website: https://deynekin.com

.GITHUB
    https://github.com/paulmann/Print-Spooler-Queue-Cleanup

.NOTES
    Version: 3.0 Professional Cross-Version Edition
    Requirements:
    - PowerShell 5.1 or later
    - Administrator privileges required for local execution
    - Windows Print Spooler service must be installed on target machines
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
    [Parameter(ValueFromPipeline = $true)]
    [string[]]$ComputerName = @('localhost'),
    [string]$LogPath,
    [switch]$Force
)

# Enable strict mode for better error handling in all versions
Set-StrictMode -Version Latest

begin {
    $ScriptVersion = "3.0-Professional-CrossVersion"
    $StartTime = Get-Date
    $ScriptName = $MyInvocation.MyCommand.Name
    $ScriptPath = $MyInvocation.MyCommand.Path
    $ScriptDirectory = Split-Path -Parent $ScriptPath

    # Configuration
    $ScriptConfig = @{
        ServiceName      = 'Spooler'
        RequiredPSVersion = [Version]'5.1'
        WaitTimeSeconds  = 3
        LogPrefix        = '[Print Spooler Cleanup]'
    }

    # Statistics object
    $Global:Statistics = @{
        StartTime        = $StartTime
        EndTime          = $null
        FilesProcessed   = 0
        JobsCleared      = 0
        ErrorsEncountered = 0
        OperationResult  = 'Unknown'
    }

    if (-not $LogPath) {
        $Timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $LogPath = Join-Path -Path $ScriptDirectory -ChildPath "Clear-PrintSpoolerQueue_$Timestamp.log"
    }

    # --- Helper Functions ---
    function Write-Log {
        param([string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')][string]$Level = 'INFO')
        $Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $LogEntry = "[$Timestamp] [$Level] $Message"
        Write-Verbose $LogEntry
        try { "$LogEntry" | Out-File -FilePath $LogPath -Append -Encoding UTF8 -ErrorAction Stop }
        catch { Write-Error "Log write failed: $_" }
    }

    function Write-HeaderMessage {
        $headerLines = @(
            ("═" * 80),
            "    PowerShell Print Spooler Queue Cleanup Utility v$ScriptVersion",
            "    Professional System Administrator Tool",
            "    Author: Mikhail Deynekin",
            "    GitHub: https://github.com/mdeynekin/Print-Spooler-Queue-Cleanup",
            ("═" * 80),
            ""
        )
        foreach ($line in $headerLines) {
            Write-Host $line -ForegroundColor Cyan
        }
    }

    function Write-StatusMessage {
        param(
            [Parameter(Mandatory)]
            [string]$Message,
            [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Progress')]
            [string]$Level = 'Info'
        )
        $timestamp = Get-Date -Format 'HH:mm:ss'
        $prefix = "$($ScriptConfig.LogPrefix) [$timestamp]"
        # Map levels to icons and colors
        $iconMap = @{ 'Info' = 'ℹ'; 'Success' = '✓'; 'Warning' = '⚠'; 'Error' = '✗'; 'Progress' = '►' }
        $colorMap = @{ 'Info' = 'White'; 'Success' = 'Green'; 'Warning' = 'Yellow'; 'Error' = 'Red'; 'Progress' = 'Cyan' }
        $icon = $iconMap[$Level]
        $color = $colorMap[$Level]
        $formattedMessage = "$prefix $icon $Message"
        Write-Host $formattedMessage -ForegroundColor $color
        # Also log it
        $logLevel = switch ($Level) { 'Warning' { 'WARN' } 'Error' { 'ERROR' } default { 'INFO' } }
        Write-Log -Message $Message -Level $logLevel
    }

    function Test-AdministratorRights {
        try {
            $currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
            return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        }
        catch {
            Write-StatusMessage "Failed to check administrator rights: $($_.Exception.Message)" -Level Error
            return $false
        }
    }

    function Test-PowerShellVersion {
        $currentVersion = $PSVersionTable.PSVersion
        if ($currentVersion -lt $ScriptConfig.RequiredPSVersion) {
            Write-StatusMessage "PowerShell version $currentVersion is below required $($ScriptConfig.RequiredPSVersion)" -Level Error
            return $false
        }
        return $true
    }

    # --- Cross-Version Safe Core Functions ---
    function Invoke-ScriptBlockOnTarget {
        param(
            [string]$Computer,
            [scriptblock]$ScriptBlock,
            [object[]]$ArgumentList = @()
        )
        if ($Computer -eq 'localhost' -or $Computer -eq '.' -or $Computer -eq $env:COMPUTERNAME) {
            if ($ArgumentList.Count -eq 0) {
                return & $ScriptBlock
            } else {
                return & $ScriptBlock @ArgumentList
            }
        } else {
            $icmParams = @{
                ComputerName = $Computer
                ScriptBlock  = $ScriptBlock
                ErrorAction  = 'Stop'
            }
            if ($ArgumentList.Count -gt 0) { $icmParams['ArgumentList'] = $ArgumentList }
            return Invoke-Command @icmParams
        }
    }

    function Get-SpoolerDirectoryPath {
        param([string]$TargetComputer)
        $ScriptBlock = {
            $SpoolPath = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers" -Name "DefaultSpoolDirectory" -ErrorAction SilentlyContinue)."DefaultSpoolDirectory"
            if (-not $SpoolPath) { "$env:SystemRoot\System32\spool\PRINTERS" } else { $SpoolPath }
        }
        return Invoke-ScriptBlockOnTarget -Computer $TargetComputer -ScriptBlock $ScriptBlock
    }

    function Get-SpoolFileStatistics {
        param([string]$TargetComputer, [string]$SpoolDir)
        $ScriptBlock = {
            param($Path)
            if (-not (Test-Path $Path)) { return @{ Files=0; SHDFiles=0; SPLFiles=0; TotalSize=0 } }
            $allFiles = Get-ChildItem -Path $Path -File -ErrorAction SilentlyContinue
            $shd = ($allFiles | Where-Object Extension -eq '.SHD').Count
            $spl = ($allFiles | Where-Object Extension -eq '.SPL').Count
            $size = ($allFiles | Measure-Object -Property Length -Sum).Sum
            @{ Files=$allFiles.Count; SHDFiles=$shd; SPLFiles=$spl; TotalSize=$size }
        }
        return Invoke-ScriptBlockOnTarget -Computer $TargetComputer -ScriptBlock $ScriptBlock -ArgumentList $SpoolDir
    }

    function Stop-SpoolerServiceOnTarget {
        param([string]$TargetComputer)
        Write-StatusMessage "Stopping Print Spooler service on '$TargetComputer'..." -Level Progress
        try {
            $service = Get-Service -Name $ScriptConfig.ServiceName -ComputerName $TargetComputer -ErrorAction Stop
            if ($service.Status -eq 'Running') {
                Stop-Service -Name $ScriptConfig.ServiceName -ComputerName $TargetComputer -Force -ErrorAction Stop
                # Wait with progress bar
                $timeout = 30
                for ($i = 1; $i -le $timeout; $i++) {
                    Start-Sleep -Seconds 1
                    $service = Get-Service -Name $ScriptConfig.ServiceName -ComputerName $TargetComputer
                    if ($service.Status -eq 'Stopped') { break }
                    Write-Progress -Activity "Stopping Print Spooler on '$TargetComputer'" -Status "Waiting for service to stop..." -PercentComplete (($i / $timeout) * 100)
                }
                Write-Progress -Activity "Stopping Print Spooler" -Completed
                if ($service.Status -eq 'Stopped') {
                    Write-StatusMessage "Print Spooler service stopped successfully on '$TargetComputer'" -Level Success
                    return $true
                } else {
                    throw "Service did not stop within timeout period"
                }
            } else {
                Write-StatusMessage "Print Spooler service was already stopped on '$TargetComputer'" -Level Info
                return $true
            }
        }
        catch {
            Write-StatusMessage "Failed to stop Print Spooler service on '$TargetComputer': $($_.Exception.Message)" -Level Error
            $Global:Statistics.ErrorsEncountered++
            return $false
        }
    }

    function Clear-SpoolFilesOnTarget {
        param([string]$TargetComputer, [string]$SpoolDir)
        Write-StatusMessage "Analyzing spool directory on '$TargetComputer'..." -Level Progress
        $preStats = Get-SpoolFileStatistics -TargetComputer $TargetComputer -SpoolDir $SpoolDir
        if ($preStats.Files -eq 0) {
            Write-StatusMessage "No files found in spool directory on '$TargetComputer' - queue already clean" -Level Info
            return $true
        }
        Write-StatusMessage "Found $($preStats.Files) files ($($preStats.SHDFiles) .SHD, $($preStats.SPLFiles) .SPL) on '$TargetComputer'" -Level Info

        $ScriptBlock = {
            param($Path)
            $files = @(Get-ChildItem -Path $ScriptConfig.SpoolDirectory -File -ErrorAction SilentlyContinue)
            $results = @()
            foreach ($file in $files) {
                try {
                    Remove-Item -Path $file.FullName -Force -ErrorAction Stop
                    $results += @{ Success = $true; File = $file.Name }
                }
                catch {
                    $results += @{ Success = $false; File = $file.Name; Error = $_.Exception.Message }
                }
            }
            return $results
        }

        Write-StatusMessage "Clearing spool files on '$TargetComputer'..." -Level Progress
        $results = Invoke-ScriptBlockOnTarget -Computer $TargetComputer -ScriptBlock $ScriptBlock -ArgumentList $SpoolDir
        $successCount = ($results | Where-Object Success -eq $true).Count
        $errorCount = ($results | Where-Object Success -eq $false).Count

        if ($errorCount -gt 0) {
            foreach ($result in ($results | Where-Object Success -eq $false)) {
                Write-StatusMessage "Failed to remove file '$($result.File)': $($result.Error)" -Level Warning
            }
            $Global:Statistics.ErrorsEncountered += $errorCount
        }

        $Global:Statistics.FilesProcessed += $successCount
        $Global:Statistics.JobsCleared = [Math]::Max($preStats.SHDFiles, $preStats.SPLFiles)
        Write-StatusMessage "Successfully processed $successCount files on '$TargetComputer'" -Level Success
        return $true
    }

    function Start-SpoolerServiceOnTarget {
        param([string]$TargetComputer)
        Write-StatusMessage "Starting Print Spooler service on '$TargetComputer'..." -Level Progress
        try {
            Start-Service -Name $ScriptConfig.ServiceName -ComputerName $TargetComputer -ErrorAction Stop
            $timeout = 30
            for ($i = 1; $i -le $timeout; $i++) {
                Start-Sleep -Seconds 1
                $service = Get-Service -Name $ScriptConfig.ServiceName -ComputerName $TargetComputer
                if ($service.Status -eq 'Running') { break }
                Write-Progress -Activity "Starting Print Spooler on '$TargetComputer'" -Status "Waiting for service to start..." -PercentComplete (($i / $timeout) * 100)
            }
            Write-Progress -Activity "Starting Print Spooler" -Completed
            if ($service.Status -eq 'Running') {
                Write-StatusMessage "Print Spooler service started successfully on '$TargetComputer'" -Level Success
                return $true
            } else {
                throw "Service did not start within timeout period"
            }
        }
        catch {
            Write-StatusMessage "Failed to start Print Spooler service on '$TargetComputer': $($_.Exception.Message)" -Level Error
            $Global:Statistics.ErrorsEncountered++
            return $false
        }
    }

    function Show-ExecutionStatistics {
        $Global:Statistics.EndTime = Get-Date
        $executionTime = $Global:Statistics.EndTime - $Global:Statistics.StartTime
        Write-Host "`n" -NoNewline
        Write-Host ("═" * 80) -ForegroundColor Cyan
        Write-Host "    EXECUTION STATISTICS & SUMMARY" -ForegroundColor Cyan
        Write-Host ("═" * 80) -ForegroundColor Cyan
        $statsDisplay = @(
            @{ Label = "Execution Time"; Value = "{0:mm\:ss\.fff}" -f $executionTime; Color = "White" },
            @{ Label = "Print Jobs Cleared"; Value = $Global:Statistics.JobsCleared; Color = if ($Global:Statistics.JobsCleared -gt 0) { "Green" } else { "Yellow" } },
            @{ Label = "Files Processed"; Value = $Global:Statistics.FilesProcessed; Color = if ($Global:Statistics.FilesProcessed -gt 0) { "Green" } else { "Yellow" } },
            @{ Label = "Errors Encountered"; Value = $Global:Statistics.ErrorsEncountered; Color = if ($Global:Statistics.ErrorsEncountered -eq 0) { "Green" } else { "Red" } },
            @{ Label = "Operation Result"; Value = $Global:Statistics.OperationResult; Color = if ($Global:Statistics.OperationResult -eq "Success") { "Green" } else { "Red" } }
        )
        foreach ($stat in $statsDisplay) {
            $padding = " " * (20 - $stat.Label.Length)
            Write-Host "    $($stat.Label):$padding" -NoNewline -ForegroundColor Gray
            Write-Host $stat.Value -ForegroundColor $stat.Color
        }
        Write-Host ("═" * 80) -ForegroundColor Cyan

        if ($Global:Statistics.OperationResult -eq "Success" -and $Global:Statistics.ErrorsEncountered -eq 0) {
            Write-StatusMessage "Print spooler cleanup completed successfully!" -Level Success
        } elseif ($Global:Statistics.OperationResult -eq "Success" -and $Global:Statistics.ErrorsEncountered -gt 0) {
            Write-StatusMessage "Print spooler cleanup completed with $($Global:Statistics.ErrorsEncountered) warnings" -Level Warning
        } else {
            Write-StatusMessage "Print spooler cleanup failed - check error messages above" -Level Error
        }
    }

    # --- Initialization ---
    Write-HeaderMessage
    Write-Log "Script '$ScriptName' v$ScriptVersion started on PowerShell $($PSVersionTable.PSVersion)."
    if (-not (Test-PowerShellVersion)) { throw "PowerShell version requirements not met." }
    if ($ComputerName -contains 'localhost' -or $ComputerName -contains '.' -or $ComputerName -contains $env:COMPUTERNAME) {
        if (-not (Test-AdministratorRights)) { throw "Administrator privileges required." }
    }
}

process {
    foreach ($Computer in $ComputerName) {
        Write-StatusMessage "Starting cleanup process on computer: '$Computer'" -Level Info

        try {
            $SpoolDir = Get-SpoolerDirectoryPath -TargetComputer $Computer
            Write-Log "Spool directory on '$Computer': '$SpoolDir'"

            if (-not $Force -and -not $PSCmdlet.ShouldProcess($Computer, "Clear Print Spooler Queue")) {
                Write-StatusMessage "Operation cancelled by user for '$Computer'" -Level Info
                continue
            }

            if (-not (Stop-SpoolerServiceOnTarget -TargetComputer $Computer)) {
                throw "Failed to stop service on '$Computer'"
            }

            Start-Sleep -Seconds $ScriptConfig.WaitTimeSeconds

            if (-not (Clear-SpoolFilesOnTarget -TargetComputer $Computer -SpoolDir $SpoolDir)) {
                Write-StatusMessage "File cleanup encountered issues on '$Computer', continuing..." -Level Warning
            }

            if (-not (Start-SpoolerServiceOnTarget -TargetComputer $Computer)) {
                throw "Failed to restart service on '$Computer'"
            }

            Write-StatusMessage "Cleanup completed successfully on '$Computer'" -Level Success
        }
        catch {
            $ErrorMessage = "Critical error on '$Computer': $($_.Exception.Message)"
            Write-StatusMessage $ErrorMessage -Level Error
            $Global:Statistics.ErrorsEncountered++
            $Global:Statistics.OperationResult = "Failed"
            # Attempt recovery
            try {
                Write-StatusMessage "Attempting service recovery on '$Computer'..." -Level Warning
                Start-Service -Name $ScriptConfig.ServiceName -ComputerName $Computer -ErrorAction SilentlyContinue
            }
            catch {
                Write-StatusMessage "Service recovery failed on '$Computer': $($_.Exception.Message)" -Level Error
            }
        }
    }
}

end {
    # If we processed only one computer, show detailed stats
    if ($ComputerName.Count -eq 1) {
        if ($Global:Statistics.ErrorsEncountered -eq 0) { $Global:Statistics.OperationResult = "Success" }
        Show-ExecutionStatistics
    }
    else {
        # For multiple computers, just log completion
        Write-Log "Bulk operation completed for $($ComputerName.Count) computers."
    }
    $exitCode = if ($Global:Statistics.OperationResult -eq "Success") { 0 } else { 1 }
    exit $exitCode
}

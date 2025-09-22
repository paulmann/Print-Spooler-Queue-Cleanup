<#
.SYNOPSIS
    Forcibly clears the Windows Print Spooler queue.

.DESCRIPTION
    This script checks for administrative rights, stops the Print Spooler service,
    deletes all files in the spooler directory, and restarts the service. It includes
    robust error handling and verbose logging for Senior-level system administration.

.AUTHOR
    Mikhail Deynekin

.NOTES
    - Requires PowerShell 5.1 or later
    - Must be run as Administrator

.EXAMPLE
    PS C:\> .\Clear-PrintSpoolerQueue.ps1
#>

# Verify script is running with administrative privileges
if (-not ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltinRole]::Administrator)) {
    Write-Error "This script must be run as Administrator."
    Exit 1
}

# Define the spooler service name and spool directory path
$serviceName = 'Spooler'
$spoolDir     = Join-Path $env:SystemRoot 'System32\spool\PRINTERS'

# Stop the Print Spooler service
Write-Verbose "Stopping the Print Spooler service..."
try {
    Stop-Service -Name $serviceName -Force -ErrorAction Stop
    Write-Verbose "Service stopped successfully."
} catch {
    Write-Error "Failed to stop service: $_"
    Exit 1
}

# Wait to ensure the service has fully stopped
Start-Sleep -Seconds 3

# Remove all spool files (*.SPL and *.SHD)
Write-Verbose "Removing spool files from $spoolDir..."
try {
    Remove-Item -Path "$spoolDir\*" -Force -Recurse -ErrorAction Stop
    Write-Verbose "Spool directory cleaned."
} catch {
    Write-Error "Failed to remove spool files: $_"
    # Attempt to restart service even if cleanup fails
}

# Restart the Print Spooler service
Write-Verbose "Restarting the Print Spooler service..."
try {
    Start-Service -Name $serviceName -ErrorAction Stop
    Write-Verbose "Service restarted successfully."
} catch {
    Write-Error "Failed to restart service: $_"
    Exit 1
}

Write-Output "Print spooler queue has been forcibly cleared and service restarted."

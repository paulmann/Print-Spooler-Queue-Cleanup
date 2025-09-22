# Print Spooler Queue Cleanup Tools
A PowerShell script to forcibly clear the print spooler queue on Windows systems. It verifies administrative privileges, stops the Print Spooler service, removes spool files, and restarts the service. Designed for Senior System Administrators to streamline print queue maintenance.

### Table of Contents
- [Overview](#overview)  
- [Problem Statement](#problem-statement)  
- [Solution](#solution)  
- [Scripts](#scripts)  
  - [Batch Script: Clear-PrintSpoolerQueue.bat](#batch-script-clear-printspoolerqueuebat)  
  - [PowerShell Script: Clear-PrintSpoolerQueue.ps1](#powershell-script-clear-printspoolerqueueps1)  
- [Prerequisites](#prerequisites)  
- [Installation](#installation)  
- [Usage](#usage)  
  - [Running the Batch Script](#running-the-batch-script)  
  - [Running the PowerShell Script](#running-the-powershell-script)  
- [Troubleshooting Execution Policy Errors](#troubleshooting-execution-policy-errors)  
- [Automation and Scheduling](#automation-and-scheduling)  
- [Script Details](#script-details)  
- [Error Handling & Logging](#error-handling--logging)  
- [Author](#author)  
- [License](#license)  

***

## Overview  
This repository contains two professional-grade tools for **forcible clearing** of Windows Print Spooler queues across all installed printers. Both scripts verify administrative privileges, stop the Print Spooler service, remove all spool files, and restart the service—delivering a consistent, reliable solution for stuck or corrupted print jobs.

## Problem Statement  
When print jobs become **stuck** or **corrupted**, the Print Spooler may hang, blocking all subsequent print operations. Manual cleanup—stopping the service, deleting files in the spool directory, and restarting—can be tedious, error-prone, and hard to standardize in enterprise environments.

## Solution  
These scripts automate the entire process:
1. **Privilege Check**: Ensures administrative rights.  
2. **Service Control**: Stops the Spooler service.  
3. **File Cleanup**: Deletes `*.SPL` and `*.SHD` files from the spool directory.  
4. **Service Restart**: Restarts the Spooler to resume printing.

Consistent automation reduces downtime, minimizes support calls, and can be integrated into remote management or scheduled maintenance.

***

## Scripts

### Batch Script: Clear-PrintSpoolerQueue.bat  
A lightweight, native Windows batch file for rapid deployment on servers and workstations without PowerShell dependencies.  
📄 [View on GitHub](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/blob/main/Clear-PrintSpoolerQueue.bat)

Key features:
- Admin privilege detection via `cacls.exe`.  
- Silent service control using `net stop` / `net start`.  
- Forced deletion of spool files (`*.spl`, `*.shd`).  
- Clear console feedback with success/error codes.

***

### PowerShell Script: Clear-PrintSpoolerQueue.ps1  
A robust PowerShell solution leveraging advanced error handling and verbose logging for Senior System Administrators.  
📄 [View on GitHub](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/blob/main/Clear-PrintSpoolerQueue.ps1)

Key features:
- WindowsPrincipal check for elevated context.  
- Execution policy verification with user guidance.  
- `Stop-Service` / `Start-Service` with `-Force` and `-ErrorAction`.  
- Recursive `Remove-Item` cleanup with error trapping.  
- Verbose and error messages for auditability.

***

## Prerequisites  
- **Operating System:** Windows 7 / 8 / 10 / 11 or Server 2008 R2 and newer  
- **Permissions:** Must run as **Administrator**  
- **PowerShell Version:** 5.1 or later  

***

## Installation  
1. Clone or download this repository.  
2. Place both scripts in a suitable directory (e.g., `C:\Scripts\PrintCleanup`).  
3. (Optional) Unblock downloaded scripts:  
   ```powershell
   Unblock-File C:\Scripts\PrintCleanup\*.ps1
   ```

***

## Usage

### Running the Batch Script  
1. Right-click **Clear-PrintSpoolerQueue.bat** → **Run as administrator**.  
2. Observe console messages for status and completion.  

### Running the PowerShell Script  
1. Open **PowerShell** as Administrator.  
2. Navigate to script directory:
   ```powershell
   cd C:\Scripts\PrintCleanup
   ```
3. Execute:
   ```powershell
   .\Clear-PrintSpoolerQueue.ps1 -Verbose
   ```
4. Review verbose output for detailed steps.

***

## Troubleshooting Execution Policy Errors  
If you encounter the error when launching the PS1 script:
```
File Clear-PrintSpoolerQueue.ps1 cannot be loaded because running scripts is disabled on this system. 
For more information, see about_Execution_Policies at https:/go.microsoft.com/fwlink/?LinkID=135170.
```
a Senior System Administrator should perform the following:

1. Open PowerShell as Administrator.  
2. Permit script execution for the current user:
   ```powershell
   Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted
   ```
3. Re-run the cleanup script.

The PowerShell script itself includes an **execution policy check** at startup. If the policy is too restrictive, it will emit a clear error directing you to adjust the execution policy before attempting cleanup.

***

## Automation and Scheduling  
To automate cleanup during off-hours, configure Task Scheduler:

1. Open **Task Scheduler** → **Create Task**.  
2. Under **General**, select **Run with highest privileges**.  
3. Under **Triggers**, set desired schedule (e.g., daily at 03:00).  
4. Under **Actions**, add:
   - **Program/script**: `powershell.exe`  
   - **Arguments**:  
     ```text
     -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\PrintCleanup\Clear-PrintSpoolerQueue.ps1" -Verbose
     ```
5. Save and enable the task.

For the batch file, point **Program/script** to the `.bat` path.

***

## Script Details

### Clear-PrintSpoolerQueue.bat
```bat
@echo off
:: ----------------------------------------------------------------------------
:: Forcibly clears the Windows Print Spooler queue.
:: - Checks for Administrator
:: - Stops Spooler service
:: - Deletes *.spl and *.shd
:: - Restarts Spooler service
:: Author: Mikhail Deynekin
:: ----------------------------------------------------------------------------

>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo ERROR: Administrator privileges required.
    pause
    exit /b 1
)

net stop Spooler >nul 2>&1 || exit /b 1
timeout /t 3 /nobreak >nul
del /F /Q "%SystemRoot%\System32\spool\PRINTERS\*.spl" >nul 2>&1
del /F /Q "%SystemRoot%\System32\spool\PRINTERS\*.shd" >nul 2>&1
net start Spooler >nul 2>&1 || exit /b 1

echo Cleanup complete.
exit /b 0
```

### Clear-PrintSpoolerQueue.ps1
```powershell
<#
.SYNOPSIS
    Forcibly clears the Windows Print Spooler queue.

.DESCRIPTION
    Checks for Administrator, verifies execution policy, stops Spooler,
    deletes spool files, restarts service.

.AUTHOR
    Mikhail Deynekin

.NOTES
    Requires PowerShell 5.1+, run as Admin.
#>

# Verify administrative privileges
if (-not ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltinRole]::Administrator)) {
    Write-Error "Administrator privileges required."
    Exit 1
}

# Verify execution policy
$policy = Get-ExecutionPolicy -Scope CurrentUser
if ($policy -in @('Restricted','Undefined')) {
    Write-Error "Current execution policy '$policy' blocks script execution. Run:" `
        + "`n  Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted" 
    Exit 1
}

$serviceName = 'Spooler'
$spoolDir     = "$env:SystemRoot\System32\spool\PRINTERS"

# Stop the Print Spooler service
Stop-Service -Name $serviceName -Force -ErrorAction Stop
Start-Sleep -Seconds 3

# Remove all spool files
Remove-Item -Path "$spoolDir\*" -Force -Recurse -ErrorAction SilentlyContinue

# Restart the Print Spooler service
Start-Service -Name $serviceName -ErrorAction Stop

Write-Output "Print Spooler queue has been forcibly cleared and service restarted."
```

***

## Error Handling & Logging  
- **Critical failures** (service stop/start errors or execution policy blocks) terminate the script with a clear error message and non-zero exit code.  
- **Non-fatal errors** during file removal are logged, but the service restart is still attempted.  
- Use the `-Verbose` switch in PowerShell to receive detailed operational output.

***

## Author  
**Mikhail Deynekin**  
Senior System Administrator & Developer  
- Email: mid1977@gmail.com  
- Website: https://deynekin.com  

***

## License  
This project is licensed under the **MIT License**. See [LICENSE](LICENSE) for details.

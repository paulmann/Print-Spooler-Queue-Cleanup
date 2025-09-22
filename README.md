# Print-Spooler-Queue-Cleanup
A PowerShell script to forcibly clear the print spooler queue on Windows systems. It verifies administrative privileges, stops the Print Spooler service, removes spool files, and restarts the service. Designed for Senior System Administrators to streamline print queue maintenance.

# Clear-PrintSpoolerQueue PowerShell Script

## Table of Contents
- [Overview](#overview)  
- [Problem Statement](#problem-statement)  
- [Solution](#solution)  
- [Prerequisites](#prerequisites)  
- [Installation](#installation)  
- [Usage](#usage)  
- [Script Details](#script-details)  
- [Error Handling & Logging](#error-handling--logging)  
- [Author](#author)  
- [License](#license)  

## Overview  
`Clear-PrintSpoolerQueue.ps1` is a professional-grade PowerShell script tailored for Senior System Administrators. It automates the **forcible clearing** of Windows Print Spooler queues across all installed printers. The script ensures reliable cleanup of hung or stuck print jobs, minimizing downtime and user support tickets.

## Problem Statement  
Print jobs can become **stuck** or **corrupted**, causing the Print Spooler service to hang and preventing all subsequent jobs from processing. Manual intervention—navigating through Services, stopping the spooler, deleting files, and restarting the service—can be time-consuming and error-prone, especially in large environments or during high-support-load periods.

## Solution  
`Clear-PrintSpoolerQueue.ps1` addresses this by:
- **Verifying administrative privileges**  
- **Stopping** the Print Spooler service (Spooler)  
- **Deleting** all spool files (`*.SPL`, `*.SHD`) in the spool directory  
- **Restarting** the service to resume printing operations  

This scripted approach ensures consistent results, reduces human error, and can be integrated into scheduled maintenance tasks or remote automation frameworks.

## Prerequisites  
- **Operating System:** Windows 7, 8, 10, 11, or Server 2008 R2 and newer  
- **PowerShell Version:** 5.1 or later  
- **Permissions:** Must be executed with **Administrator** privileges  

## Installation  
1. Clone or download this repository to your local system or management server.  
2. Place `Clear-PrintSpoolerQueue.ps1` in a directory of your choice (e.g., `C:\Scripts\`).  
3. (Optional) Unblock the script if downloaded from the internet:  
   ```powershell
   Unblock-File -Path 'C:\Scripts\Clear-PrintSpoolerQueue.ps1'
   ```

## Usage  
1. Open **PowerShell** as Administrator.  
2. Navigate to the script directory:
   ```powershell
   cd C:\Scripts
   ```
3. Run the script:
   ```powershell
   .\Clear-PrintSpoolerQueue.ps1
   ```
4. Observe verbose output for status of each step.  

### Integrating with Task Scheduler  
To automate cleanup during off-hours, create a scheduled task:
1. Open **Task Scheduler**.  
2. Create a new task with **highest privileges**.  
3. Set trigger (e.g., daily at 03:00).  
4. Action:
   - Program/script: `powershell.exe`
   - Arguments: `-NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\Clear-PrintSpoolerQueue.ps1" -Verbose`  

## Script Details  
```powershell
<#
.SYNOPSIS
    Forcibly clears the Windows Print Spooler queue.

.DESCRIPTION
    Checks for administrative rights, stops the Print Spooler service,
    deletes all files in the spool directory, and restarts the service.

.PARAMETER None
    No parameters. All configuration is internal.

.NOTES
    Author: Mikhail Deynekin
    Requires PowerShell 5.1 or later
    Run As Administrator
#>
```

- **Privilege Check**: Ensures the script exits if not run as Administrator.  
- **Service Control**: Uses `Stop-Service` and `Start-Service` with error handling.  
- **File Cleanup**: Employs `Remove-Item` to delete spool files in `%SystemRoot%\System32\spool\PRINTERS`.  
- **Verbose Logging**: All major operations emit verbose messages for auditability.

## Error Handling & Logging  
- The script **terminates** on critical failures (unable to stop or start the service).  
- Non-fatal errors during file deletion will be reported but the service restart is still attempted.  
- Use the `-Verbose` switch to view detailed processing steps.  

## Author  
**Mikhail Deynekin**  
- Email: mid1977@gmail.com  
- Website: https://deynekin.com  

## License  
This project is licensed under the **MIT License**. See [LICENSE](LICENSE) for details.

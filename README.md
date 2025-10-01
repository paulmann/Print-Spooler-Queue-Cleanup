# 🖨️ Print Spooler Queue Cleanup v3.0 Professional

[![License: Public Domain](https://img.shields.io/badge/License-Public%20Domain-brightgreen.svg)](https://creativecommons.org/publicdomain/zero/1.0/)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://docs.microsoft.com/powershell/)
[![Batch](https://img.shields.io/badge/Batch-CMD-green.svg)](https://docs.microsoft.com/windows-server/administration/windows-commands/windows-commands)
[![Windows](https://img.shields.io/badge/Windows-7%7C8%7C10%7C11%7CServer-0078d4.svg)](https://www.microsoft.com/windows)

> **Enterprise-grade PowerShell and Batch script suite for forcibly clearing Windows Print Spooler queues with comprehensive logging, error handling, and remote management capabilities**

Print Spooler Queue Cleanup is a professional system administration solution designed to resolve stuck or corrupted print jobs by safely managing the Windows Print Spooler service, removing problematic spool files, and ensuring service recovery. Perfect for Senior System Administrators managing Windows environments at scale.

## 🚀 Quick Start

```powershell
# Download and run batch script (simple approach)
# Right-click Clear-PrintSpoolerQueue.bat → Run as Administrator

# Or clone the repository
git clone https://github.com/paulmann/Print-Spooler-Queue-Cleanup.git
cd Print-Spooler-Queue-Cleanup

# Run batch script (v3.0)
Clear-PrintSpoolerQueue.bat

# Run PowerShell script with detailed output (v3.0)
powershell.exe -ExecutionPolicy Bypass -File "Clear-PrintSpoolerQueue.ps1" -Verbose

# Remote execution on multiple computers
.\Clear-PrintSpoolerQueue.ps1 -ComputerName "SERVER01","PC-FINANCE" -Force

# Schedule automated cleanup
schtasks /create /tn "Print Spooler Cleanup" /tr "C:\Scripts\Clear-PrintSpoolerQueue.bat" /sc daily /st 03:00 /ru SYSTEM
```

## 📋 Table of Contents

- [🚨 Why Print Spooler Queue Cleanup?](#-why-print-spooler-queue-cleanup)
  - [The Hidden Problem](#the-hidden-problem)
  - [Real-World Impact](#real-world-impact)
- [✨ Key Features v3.0](#-key-features-v30)
  - [🛡️ Enterprise-Grade Safety](#️-enterprise-grade-safety)
  - [🎯 Intelligent Processing](#-intelligent-processing)
  - [📊 Advanced Logging & Statistics](#-advanced-logging--statistics)
  - [🔄 Automation & Remote Management](#-automation--remote-management)
- [📦 Installation & Usage](#-installation--usage)
  - [System Requirements](#system-requirements)
  - [Installation Options](#installation-options)
  - [Command Line Usage](#command-line-usage)
  - [Usage Examples](#usage-examples)
- [🏗️ Advanced Features](#️-advanced-features)
  - [Script Comparison v3.0](#script-comparison-v30)
  - [Error Handling & Recovery](#error-handling--recovery)
  - [Cross-Version Compatibility](#cross-version-compatibility)
- [🔗 DevOps Integration](#-devops-integration)
  - [Task Scheduler Integration](#task-scheduler-integration)
  - [Group Policy Deployment](#group-policy-deployment)
  - [Remote Management](#remote-management)
- [🏢 Enterprise Usage](#-enterprise-usage)
  - [Deployment Strategies](#deployment-strategies)
  - [Monitoring & Reporting](#monitoring--reporting)
  - [Best Practices](#best-practices)
- [🔍 Troubleshooting](#-troubleshooting)
- [🤝 Contributing](#-contributing)
- [📄 License](#-license)
- [👨‍💻 Author & Support](#-author--support)

## 🚨 Why Print Spooler Queue Cleanup?

### The Hidden Problem

Print jobs can become **stuck** or **corrupted** in the Windows Print Spooler, causing:

```powershell
# Typical symptoms of stuck print spooler
PS C:\> Get-Service Spooler
Status: Running but no documents print

PS C:\> Get-Printer | Get-PrintJob
# Shows jobs stuck in queue that cannot be cancelled
```

```cmd
# Manual cleanup is tedious and error-prone
net stop spooler
del /F /Q "%SystemRoot%\System32\spool\PRINTERS\*.*"
net start spooler
# What if service fails to stop? What about permissions?
```

### Real-World Impact

- **🖨️ Print Queue Blockages**: Jobs stuck indefinitely, blocking all subsequent printing
- **⚠️ Service Hang**: Print Spooler service becomes unresponsive requiring manual intervention
- **🏢 Enterprise Downtime**: Multiple users affected, increased support ticket volume
- **⏰ Time Waste**: Manual cleanup takes time and is prone to human error
- **🔄 Recurring Issues**: Without automation, problems repeat frequently

## ✨ Key Features v3.0

### 🛡️ **Enterprise-Grade Safety**

- **Advanced Administrative Privilege Verification**: Multiple validation methods across Windows versions
- **Service State Management**: Comprehensive service dependency handling with timeout controls
- **File System Protection**: Safe deletion with detailed file tracking and recovery mechanisms
- **Cross-Platform Compatibility**: Support for Windows 7/8/10/11 and all Server versions
- **Legacy Mode Support**: Automatic detection and adaptation for older Windows versions

### 🎯 **Intelligent Processing**

- **Dynamic Spool Path Detection**: Registry-based path discovery for non-standard configurations
- **Comprehensive File Analysis**: Pre and post-cleanup statistics with size calculations
- **Service Dependency Management**: Proper handling of dependent services and recovery
- **Timeout Management**: Built-in delays and progress monitoring for all operations
- **Multi-Computer Support**: Bulk operations across multiple remote systems

### 📊 **Advanced Logging & Statistics**

- **Professional Console Output**: Color-coded status messages with Unicode symbols and timestamps
- **Detailed File Logging**: Timestamped log files with structured error reporting
- **Real-Time Statistics**: Live tracking of files processed, jobs cleared, and execution time
- **Performance Metrics**: Detailed operation summaries with success/failure rates
- **Event Log Integration**: Optional Windows Event Log integration for enterprise monitoring

### 🔄 **Automation & Remote Management**

- **Silent Operation Modes**: Unattended execution with Force parameter support
- **Remote Execution**: Native support for multiple computer targeting
- **Task Scheduler Ready**: Optimized for scheduled maintenance operations
- **WhatIf Support**: Safe preview mode for testing before execution
- **Bulk Management**: Efficient processing of multiple systems simultaneously

## 📦 Installation & Usage

### System Requirements

- **Operating System**: Windows 7/8/10/11, Windows Server 2008 R2+
- **Permissions**: Administrator privileges required for local execution
- **PowerShell Version**: 5.1 or later (for .ps1 script)
- **Dependencies**: Native Windows utilities only - no external modules required
- **Network Access**: WinRM enabled for remote operations (optional)

### Installation Options

#### Option 1: Direct Download

```powershell
# Download batch script
Invoke-WebRequest -Uri "https://github.com/paulmann/Print-Spooler-Queue-Cleanup/raw/main/Clear-PrintSpoolerQueue.bat" -OutFile "Clear-PrintSpoolerQueue.bat"

# Download PowerShell script
Invoke-WebRequest -Uri "https://github.com/paulmann/Print-Spooler-Queue-Cleanup/raw/main/Clear-PrintSpoolerQueue.ps1" -OutFile "Clear-PrintSpoolerQueue.ps1"

# Unblock PowerShell script
Unblock-File Clear-PrintSpoolerQueue.ps1
```

#### Option 2: Git Clone

```bash
git clone https://github.com/paulmann/Print-Spooler-Queue-Cleanup.git
cd Print-Spooler-Queue-Cleanup
```

#### Option 3: Enterprise Deployment

```powershell
# Copy to shared network location
$NetworkPath = "\\server\share\Scripts\PrintCleanup"
Copy-Item "*.bat", "*.ps1" -Destination $NetworkPath

# Deploy via Group Policy or SCCM
# Point to: \\server\share\Scripts\PrintCleanup\Clear-PrintSpoolerQueue.bat
```

### Command Line Usage

| Script | Version | Execution Method | Best For |
|--------|---------|------------------|----------|
| **Clear-PrintSpoolerQueue.bat** | v3.0 | Right-click → Run as Administrator | Quick manual cleanup, legacy compatibility |
| **Clear-PrintSpoolerQueue.ps1** | v3.0 | PowerShell with parameters | Advanced logging, remote management, automation |

### Usage Examples

#### Basic Execution

```cmd
# Batch script v3.0 - Enhanced with logging and statistics
Clear-PrintSpoolerQueue.bat
```

```powershell
# PowerShell script v3.0 - Professional Cross-Version Edition
.\Clear-PrintSpoolerQueue.ps1

# PowerShell with verbose output and custom log path
.\Clear-PrintSpoolerQueue.ps1 -Verbose -LogPath "C:\Admin\Logs\SpoolerCleanup.log"

# Test mode - see what would be done without making changes
.\Clear-PrintSpoolerQueue.ps1 -WhatIf
```

#### Remote Execution (New in v3.0)

```powershell
# Execute on single remote computer
.\Clear-PrintSpoolerQueue.ps1 -ComputerName "PrintServer01" -Verbose

# Execute on multiple remote computers
.\Clear-PrintSpoolerQueue.ps1 -ComputerName "SERVER01","PC-FINANCE","WORKSTATION-05" -Force

# Bulk execution with pipeline input
"SERVER01","SERVER02","SERVER03" | .\Clear-PrintSpoolerQueue.ps1 -Force -Verbose
```

#### Advanced Automation

```powershell
# Execute via Invoke-Command for enterprise scenarios
Invoke-Command -ComputerName "PrintServer01" -ScriptBlock {
    & "C:\Scripts\Clear-PrintSpoolerQueue.ps1" -Force
}

# Scheduled execution with comprehensive logging
$ScriptPath = "C:\Scripts\Clear-PrintSpoolerQueue.ps1"
$LogPath = "C:\Logs\SpoolerCleanup_$(Get-Date -Format 'yyyyMMdd').log"
& $ScriptPath -Force -LogPath $LogPath -Verbose
```

#### Scheduled Execution

```cmd
# Create daily cleanup task for Batch script
schtasks /create /tn "Daily Print Cleanup v3.0" /tr "C:\Scripts\Clear-PrintSpoolerQueue.bat" /sc daily /st 03:00 /ru SYSTEM /rl HIGHEST

# Create advanced PowerShell task with logging
schtasks /create /tn "Advanced Print Cleanup v3.0" ^
 /tr "powershell.exe -NoProfile -ExecutionPolicy Bypass -File 'C:\Scripts\Clear-PrintSpoolerQueue.ps1' -Force -Verbose" ^
 /sc weekly /d SUN /st 02:00 /ru SYSTEM /rl HIGHEST
```

## 🏗️ Advanced Features

### Script Comparison v3.0

| Feature | Batch Script v3.0 | PowerShell Script v3.0 |
|---------|-------------------|------------------------|
| **Admin Check** | `net session` method + legacy support | `WindowsPrincipal` + fallback methods |
| **Service Control** | `net stop/start` with error handling | `Stop-Service/Start-Service` with progress |
| **File Management** | `del /q` with size calculation | `Remove-Item` with detailed tracking |
| **Error Handling** | Function-based with logging | Advanced try/catch with recovery |
| **Logging** | File-based with timestamps | Multi-level logging + console output |
| **Remote Support** | Local only | Native multi-computer support |
| **Statistics** | Basic file count and size | Comprehensive execution metrics |
| **Progress Tracking** | Console messages | Progress bars and real-time updates |
| **Dependencies** | Native Windows only | PowerShell 5.1+ (cross-version compatible) |

### Error Handling & Recovery

The v3.0 scripts provide **comprehensive error handling** with automatic recovery:

```powershell
# PowerShell v3.0 error handling example
try {
    # Stop service with dependency management
    Stop-Service -Name 'Spooler' -Force -ErrorAction Stop
    
    # Wait with progress indication
    $timeout = 30
    for ($i = 1; $i -le $timeout; $i++) {
        Start-Sleep -Seconds 1
        $service = Get-Service -Name 'Spooler'
        if ($service.Status -eq 'Stopped') { break }
        Write-Progress -Activity "Stopping Print Spooler" -Status "Waiting..." -PercentComplete (($i / $timeout) * 100)
    }
    
    # Safe file cleanup with individual file tracking
    $results = @()
    Get-ChildItem -Path $SpoolPath -File | ForEach-Object {
        try {
            Remove-Item -Path $_.FullName -Force -ErrorAction Stop
            $results += @{ Success = $true; File = $_.Name }
        } catch {
            $results += @{ Success = $false; File = $_.Name; Error = $_.Exception.Message }
        }
    }
    
    # Restart service with verification
    Start-Service -Name 'Spooler' -ErrorAction Stop
    Write-Output "Print Spooler queue cleared successfully. Files processed: $(($results | Where-Object Success -eq $true).Count)"
} 
catch {
    Write-Error "Critical error during cleanup: $($_.Exception.Message)"
    
    # Attempt service recovery
    try {
        Write-Warning "Attempting service recovery..."
        Start-Service -Name 'Spooler' -ErrorAction SilentlyContinue
    } catch {
        Write-Error "Service recovery failed: $($_.Exception.Message)"
    }
    
    Exit 1
}
```

**Recovery Features v3.0:**

- 🔄 **Intelligent Service Recovery**: Multi-stage recovery with dependency handling
- 📝 **Granular Error Reporting**: Individual file processing results with specific error messages
- 🛡️ **Safe Fallback Mechanisms**: Service restart attempted even on partial failures
- 🚨 **Structured Exit Codes**: Standardized return codes for automation integration
- 📊 **Error Statistics**: Detailed tracking of error types and frequency

### Cross-Version Compatibility

```powershell
# PowerShell version detection and adaptation
$PSVersion = $PSVersionTable.PSVersion
if ($PSVersion -lt [Version]'5.1') {
    Write-Error "PowerShell 5.1 or later required. Current version: $PSVersion"
    exit 1
}

# Legacy Windows support in Batch script
for /f "tokens=4-5 delims=. " %%i in ('ver') do set "OS_MAJOR_VER=%%i"
if %OS_MAJOR_VER% LSS 6 (
    set "LEGACY_MODE=1"
    echo Legacy Windows detected - using compatibility mode
) else (
    set "LEGACY_MODE=0"
)
```

## 🔗 DevOps Integration

### Task Scheduler Integration

#### GUI Method (Enhanced v3.0)

1. Open **Task Scheduler** → **Create Task**
2. **General Tab**: 
   - Name: "Print Spooler Cleanup v3.0"
   - Select "Run with highest privileges"
   - Configure for your Windows version
3. **Triggers Tab**: Set desired schedule (daily recommended)
4. **Actions Tab**: 
   - Program: `powershell.exe`
   - Arguments: `-NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\Clear-PrintSpoolerQueue.ps1" -Force`
5. **Settings Tab**: Enable appropriate task behaviors

#### PowerShell-Based Task Creation

```powershell
# Create advanced scheduled task with PowerShell
$TaskName = "Print Spooler Cleanup v3.0 Professional"
$ScriptPath = "C:\Scripts\Clear-PrintSpoolerQueue.ps1"

$Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`" -Force -Verbose"
$Trigger = New-ScheduledTaskTrigger -Daily -At "03:00"
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
$Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principal
```

### Group Policy Deployment

#### Advanced GPO Configuration

```xml
<!-- Enhanced Group Policy Preferences for v3.0 -->
<ScheduledTask clsid="{CC63F200-7309-4ba0-B154-A71CD118DBCC}">
  <Properties action="C" name="Print Spooler Cleanup v3.0" runAs="SYSTEM" logonType="S4U"/>
  <Task version="1.3">
    <RegistrationInfo>
      <Description>Professional Print Spooler maintenance utility v3.0</Description>
      <Author>IT Department</Author>
    </RegistrationInfo>
    <Triggers>
      <CalendarTrigger>
        <StartBoundary>2025-01-01T03:00:00</StartBoundary>
        <ScheduleByDay>
          <DaysInterval>1</DaysInterval>
        </ScheduleByDay>
      </CalendarTrigger>
    </Triggers>
    <Actions>
      <Exec>
        <Command>powershell.exe</Command>
        <Arguments>-NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\Clear-PrintSpoolerQueue.ps1" -Force</Arguments>
      </Exec>
    </Actions>
    <Settings>
      <AllowStartOnDemand>true</AllowStartOnDemand>
      <AllowHardTerminate>false</AllowHardTerminate>
      <StartWhenAvailable>true</StartWhenAvailable>
    </Settings>
  </Task>
</ScheduledTask>
```

### Remote Management (New in v3.0)

```powershell
# Enterprise-scale remote management
$Computers = Get-ADComputer -Filter {OperatingSystem -like "*Windows*"} | Select-Object -ExpandProperty Name
$ScriptPath = "C:\Scripts\Clear-PrintSpoolerQueue.ps1"

# Parallel execution with job management
$Jobs = foreach ($Computer in $Computers) {
    Start-Job -ScriptBlock {
        param($ComputerName, $ScriptPath)
        try {
            Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                param($Path)
                & $Path -Force
            } -ArgumentList $ScriptPath
            [PSCustomObject]@{
                Computer = $ComputerName
                Status = "Success"
                Message = "Cleanup completed"
            }
        } catch {
            [PSCustomObject]@{
                Computer = $ComputerName
                Status = "Failed" 
                Message = $_.Exception.Message
            }
        }
    } -ArgumentList $Computer, $ScriptPath
}

# Monitor and collect results
$Results = $Jobs | Wait-Job | Receive-Job
$Jobs | Remove-Job

# Generate summary report
$SuccessCount = ($Results | Where-Object Status -eq "Success").Count
$FailCount = ($Results | Where-Object Status -eq "Failed").Count

Write-Host "Bulk cleanup completed: $SuccessCount successful, $FailCount failed" -ForegroundColor Green
```

## 🏢 Enterprise Usage

### Deployment Strategies

#### Small Office (1-50 computers)

```powershell
# Simple network share deployment with v3.0 features
$SharePath = "\\SERVER\Scripts$\PrintCleanup"
New-Item -Path $SharePath -ItemType Directory -Force
Copy-Item "Clear-PrintSpoolerQueue.*" -Destination $SharePath

# Create GPO startup script pointing to network share
# Computer Configuration > Policies > Windows Settings > Scripts > Startup
```

#### Medium Enterprise (50-500 computers)

```powershell
# SCCM/Intune deployment package
$PackageName = "Print Spooler Cleanup v3.0"
$SourcePath = "\\SERVER\Software$\PrintCleanup"

# Create application deployment with detection rules
# Detection: File exists C:\Scripts\Clear-PrintSpoolerQueue.ps1 version 3.0
# Install command: xcopy /Y /E "Clear-PrintSpoolerQueue.*" "C:\Scripts\"
# Schedule: Daily maintenance window
```

#### Large Enterprise (500+ computers)

```powershell
# PowerShell DSC configuration for v3.0
Configuration PrintSpoolerMaintenance {
    Node $ComputerName {
        File PrintCleanupScript {
            DestinationPath = "C:\Scripts\Clear-PrintSpoolerQueue.ps1"
            SourcePath = "\\SERVER\DSC$\Scripts\Clear-PrintSpoolerQueue.ps1"
            Ensure = "Present"
            Type = "File"
            Checksum = "SHA-256"
        }
        
        ScheduledTask PrintCleanupTask {
            TaskName = "Print Spooler Cleanup v3.0"
            ActionExecutable = "powershell.exe"
            ActionArguments = "-NoProfile -ExecutionPolicy Bypass -File C:\Scripts\Clear-PrintSpoolerQueue.ps1 -Force"
            ScheduleType = "Daily"
            StartTime = "03:00"
            RunLevel = "Highest"
            ExecuteAsCredential = $SystemCredential
        }
    }
}
```

### Monitoring & Reporting

#### Advanced Event Log Integration

```powershell
# Enhanced logging for v3.0 - add to PowerShell script
function Write-EnterpriseLog {
    param(
        [string]$Message,
        [ValidateSet('Information','Warning','Error')]$EntryType = 'Information',
        [int]$EventId = 1001
    )
    
    try {
        # Create custom event source if it doesn't exist
        if (-not (Get-EventLog -LogName Application -Source "Print Spooler Cleanup v3.0" -ErrorAction SilentlyContinue)) {
            New-EventLog -LogName Application -Source "Print Spooler Cleanup v3.0"
        }
        
        Write-EventLog -LogName Application -Source "Print Spooler Cleanup v3.0" -EventId $EventId -EntryType $EntryType -Message $Message
    } catch {
        Write-Warning "Failed to write to event log: $($_.Exception.Message)"
    }
}

# Usage examples
Write-EnterpriseLog "Spooler cleanup started on $env:COMPUTERNAME" -EventId 1001
Write-EnterpriseLog "Cleanup completed: $FilesDeleted files processed" -EventId 1002
Write-EnterpriseLog "Critical error during cleanup: $ErrorMessage" -EntryType Error -EventId 1003
```

#### Performance Monitoring Dashboard

```powershell
# Create monitoring queries for enterprise dashboards
# SCOM/PRTG/Nagios integration examples

# Query cleanup success rate across domain
Get-WinEvent -FilterHashtable @{
    LogName = 'Application'
    ProviderName = 'Print Spooler Cleanup v3.0'
    StartTime = (Get-Date).AddDays(-7)
} | Group-Object Id | ForEach-Object {
    [PSCustomObject]@{
        EventType = switch ($_.Name) {
            1001 { "Cleanup Started" }
            1002 { "Cleanup Successful" }  
            1003 { "Cleanup Failed" }
        }
        Count = $_.Count
        Percentage = [math]::Round(($_.Count / $TotalEvents) * 100, 2)
    }
}

# Performance counter monitoring
Get-Counter -Counter @(
    "\Print Queue(*)\Jobs"
    "\Print Queue(*)\Jobs Spooling"
    "\Process(spoolsv)\% Processor Time"
) -SampleInterval 60 -MaxSamples 1440  # 24 hours of data
```

### Best Practices v3.0

#### Security Considerations

- ✅ **Principle of Least Privilege**: Use service accounts with minimal required permissions
- ✅ **Script Signing**: Sign PowerShell scripts in enterprise environments
- ✅ **Secure Storage**: Store scripts in protected locations with proper ACLs
- ✅ **Audit Logging**: Enable comprehensive logging for compliance requirements
- ✅ **Network Security**: Use secure channels for remote script execution

#### Operational Excellence

- 📊 **Metrics-Driven**: Track cleanup success rates, execution times, and error patterns
- 🔄 **Automated Testing**: Regular validation of script functionality across Windows versions
- 📝 **Documentation**: Maintain updated runbooks and troubleshooting procedures
- 🚨 **Alerting**: Configure alerts for cleanup failures or performance degradation
- 🔄 **Version Control**: Use Git for script versioning and change management

## 🔍 Troubleshooting

### Common Issues v3.0

#### Access Denied Errors

```powershell
# Problem: "Access is denied" when stopping service
# Solution: Verify administrative privileges and service permissions

# Check current privileges
whoami /priv | findstr SeServiceLogonRight
Get-LocalGroupMember -Group "Administrators" | Where-Object Name -eq $env:USERNAME

# Verify service permissions
$Service = Get-Service -Name Spooler
$ServiceACL = Get-Acl -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Spooler"
$ServiceACL.Access | Where-Object IdentityReference -like "*$env:USERNAME*"
```

#### Service Dependencies (Enhanced)

```powershell
# Problem: Print Spooler service won't stop due to dependencies
# Solution: Enhanced dependency management in v3.0

function Stop-ServiceWithDependencies {
    param([string]$ServiceName)
    
    $Service = Get-Service -Name $ServiceName
    $DependentServices = $Service.DependentServices | Where-Object Status -eq 'Running'
    
    if ($DependentServices) {
        Write-Host "Stopping dependent services first..." -ForegroundColor Yellow
        foreach ($DepService in $DependentServices) {
            Write-Host "  Stopping $($DepService.Name)..." -ForegroundColor Cyan
            Stop-Service -Name $DepService.Name -Force -ErrorAction Continue
        }
    }
    
    # Now stop the main service
    Stop-Service -Name $ServiceName -Force
    
    # Restart dependent services after main service restart
    if ($DependentServices) {
        Start-Service -Name $ServiceName
        foreach ($DepService in $DependentServices) {
            Write-Host "  Restarting $($DepService.Name)..." -ForegroundColor Green
            Start-Service -Name $DepService.Name -ErrorAction Continue
        }
    }
}
```

#### PowerShell Execution Policy (Advanced)

```powershell
# Problem: Execution policy blocks script in enterprise environment
# Solution: Multi-level policy management

# Check all execution policy scopes
Get-ExecutionPolicy -List

# Enterprise-safe policy configuration
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force

# For Group Policy managed environments
if ((Get-ExecutionPolicy -Scope MachinePolicy) -eq 'Restricted') {
    Write-Warning "Machine policy restricts execution. Using bypass method..."
    $ScriptPath = $MyInvocation.MyCommand.Path
    Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -File `"$ScriptPath`"" -Verb RunAs
    exit
}

# Alternative: Use encoded commands for maximum compatibility
$ScriptContent = Get-Content $MyInvocation.MyCommand.Path -Raw
$EncodedScript = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($ScriptContent))
powershell.exe -EncodedCommand $EncodedScript
```

### Advanced Debugging

```powershell
# Enhanced debugging capabilities in v3.0
function Get-SpoolerDiagnostics {
    $Diagnostics = @{
        ServiceStatus = Get-Service -Name Spooler
        SpoolDirectory = $env:SystemRoot + '\System32\spool\PRINTERS'
        SpoolFiles = @()
        PrintQueues = @()
        RecentEvents = @()
    }
    
    # Analyze spool files
    if (Test-Path $Diagnostics.SpoolDirectory) {
        $Diagnostics.SpoolFiles = Get-ChildItem -Path $Diagnostics.SpoolDirectory -File | ForEach-Object {
            @{
                Name = $_.Name
                Size = $_.Length
                Created = $_.CreationTime
                Modified = $_.LastWriteTime
                Extension = $_.Extension
            }
        }
    }
    
    # Get print queue status
    try {
        $Diagnostics.PrintQueues = Get-Printer | ForEach-Object {
            $Jobs = Get-PrintJob -PrinterName $_.Name -ErrorAction SilentlyContinue
            @{
                PrinterName = $_.Name
                Status = $_.PrinterStatus
                JobCount = $Jobs.Count
                QueuedJobs = $Jobs | Select-Object Id, JobStatus, Size, SubmittedTime
            }
        }
    } catch {
        $Diagnostics.PrintQueues = @("Error retrieving print queues: $($_.Exception.Message)")
    }
    
    # Get recent spooler events
    $Diagnostics.RecentEvents = Get-WinEvent -FilterHashtable @{
        LogName = 'System'
        ProviderName = 'Service Control Manager'
        StartTime = (Get-Date).AddHours(-24)
    } | Where-Object Message -like "*Spooler*" | Select-Object TimeCreated, LevelDisplayName, Message
    
    return $Diagnostics
}

# Usage
$Diagnostics = Get-SpoolerDiagnostics
$Diagnostics | ConvertTo-Json -Depth 4 | Out-File "SpoolerDiagnostics_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
```

## 🤝 Contributing

We welcome contributions to make this tool even better! Here's how to get involved:

### Development Environment Setup

```bash
# Clone the repository
git clone https://github.com/paulmann/Print-Spooler-Queue-Cleanup.git
cd Print-Spooler-Queue-Cleanup

# Create development branch
git checkout -b feature/your-enhancement-name

# Test environment setup
# Ensure you have:
# - Windows test machines (various versions)
# - PowerShell 5.1+ and 7+
# - Administrative privileges
# - Test printers configured
```

### Contribution Guidelines

1. **🍴 Fork** the repository
2. **🔨 Create** a feature branch (`git checkout -b feature/amazing-feature`)
3. **🧪 Test** thoroughly on multiple Windows versions and PowerShell editions
4. **📝 Document** changes in both code comments and README updates
5. **✅ Commit** with descriptive messages (`git commit -m 'Add amazing feature'`)
6. **📤 Push** to your branch (`git push origin feature/amazing-feature`)
7. **🔄 Open** a Pull Request with detailed description

### Code Standards v3.0

- ✅ **Cross-Platform Testing**: Verify compatibility with Windows 7-11 and Server versions
- ✅ **PowerShell Compatibility**: Test with PowerShell 5.1, 7.x, and Windows PowerShell
- ✅ **Error Handling**: Comprehensive error checking with graceful degradation
- ✅ **Documentation**: Clear comments, help text, and usage examples
- ✅ **Performance**: Optimize for speed and resource usage
- ✅ **Security**: Follow security best practices and avoid privileged operations where possible

### Testing Checklist

- [ ] Batch script runs on Windows 7, 10, 11, Server 2019/2022
- [ ] PowerShell script works with PS 5.1 and 7.x
- [ ] Administrative privilege detection works correctly
- [ ] Service stop/start operations handle edge cases
- [ ] File cleanup properly handles locked files
- [ ] Logging captures all relevant information
- [ ] Remote execution functions properly
- [ ] Error recovery mechanisms activate correctly

## 📄 License

This project is dedicated to the **public domain** under the CC0 1.0 Universal license.

```
To the extent possible under law, Mikhail Deynekin has waived all copyright 
and related or neighboring rights to Print Spooler Queue Cleanup.

This work is published from: Russian Federation.

You can copy, modify, distribute and perform the work, even for commercial 
purposes, all without asking permission. No attribution is required, but 
appreciated.

For more information: https://creativecommons.org/publicdomain/zero/1.0/
```

## 👨‍💻 Author & Support

**Mikhail Deynekin**  
*Senior System Administrator & Developer*

- 🌐 **Website**: [deynekin.com](https://deynekin.com)
- 📧 **Email**: mid1977@gmail.com  
- 🐙 **GitHub**: [@paulmann](https://github.com/paulmann)
- 💼 **LinkedIn**: Professional system administration expert

### Getting Help

- 📖 **Documentation**: Read this comprehensive README
- 🐛 **Bug Reports**: [Open an issue](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/issues/new?template=bug_report.md)
- 💡 **Feature Requests**: [Request features](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/issues/new?template=feature_request.md)
- 💬 **Questions**: Check [Discussions](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/discussions)
- 📚 **Enterprise Support**: Contact via email for consulting services

### Related Projects

- 🖨️ **Print-Management-Tools** - Additional Windows print management utilities
- 🖥️ **Windows-System-Administration** - Comprehensive Windows admin tool collection  
- 📜 **PowerShell-Scripts** - Professional PowerShell utility library
- 🔧 **System-Maintenance-Suite** - Complete system maintenance automation

### ⭐ Star this repository if it helped you!

**Print Spooler Queue Cleanup v3.0 Professional** - Keeping enterprise printers flowing smoothly, one cleanup at a time! 🖨️✨

---

<div align="center">

**[🐛 Report Bug](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/issues) • [💡 Request Feature](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/issues) • [📚 Documentation](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki)**

*Made with ❤️ for System Administrators worldwide*

</div>

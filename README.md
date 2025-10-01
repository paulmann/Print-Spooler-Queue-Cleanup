# Print Spooler Queue Cleanup 🖨️🧹

[![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)](https://github.com/paulmann/Print-Spooler-Queue-Cleanup)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows-lightblue.svg)]()
[![PowerShell](https://img.shields.io/badge/powershell-5.1+-orange.svg)](https://docs.microsoft.com/en-us/powershell/)

> **A professional PowerShell and Batch script suite for forcibly clearing Windows Print Spooler queues with administrative privilege verification and comprehensive error handling**

Print Spooler Queue Cleanup is a robust, enterprise-grade solution designed to resolve stuck or corrupted print jobs by automatically stopping the Print Spooler service, removing problematic spool files, and restarting the service. Perfect for Senior System Administrators managing Windows environments.

## ⚡ Quick Start

```batch
# Download and run batch script (simple approach)
# Right-click Clear-PrintSpoolerQueue.bat → Run as administrator

# Or clone the repository
git clone https://github.com/paulmann/Print-Spooler-Queue-Cleanup.git
cd Print-Spooler-Queue-Cleanup

# Run batch script
Clear-PrintSpoolerQueue.bat

# Run PowerShell script with verbose output
powershell.exe -ExecutionPolicy Bypass -File "Clear-PrintSpoolerQueue.ps1" -Verbose

# Automated execution (no user interaction)
Clear-PrintSpoolerQueue.bat

# Schedule via Task Scheduler
schtasks /create /tn "Print Spooler Cleanup" /tr "C:\Scripts\Clear-PrintSpoolerQueue.bat" /sc daily /st 03:00
```

## 📋 Table of Contents

- [🚨 Why Print Spooler Queue Cleanup?](#-why-print-spooler-queue-cleanup)
  - [The Hidden Problem](#the-hidden-problem)
  - [Real-World Impact](#real-world-impact)
- [✨ Key Features](#-key-features)
  - [🛡️ Enterprise-Grade Safety](#️-enterprise-grade-safety)
  - [🎯 Intelligent Processing](#-intelligent-processing)
  - [📊 Comprehensive Logging](#-comprehensive-logging)
  - [🔄 Automation Ready](#-automation-ready)
- [📋 Installation & Usage](#-installation--usage)
  - [System Requirements](#system-requirements)
  - [Installation Options](#installation-options)
  - [Command Line Usage](#command-line-usage)
  - [Usage Examples](#usage-examples)
- [🏗️ Advanced Features](#️-advanced-features)
  - [Script Comparison](#script-comparison)
  - [Error Handling & Recovery](#error-handling--recovery)
  - [Execution Policy Management](#execution-policy-management)
- [🔗 DevOps Integration](#-devops-integration)
  - [Task Scheduler Integration](#task-scheduler-integration)
  - [Group Policy Deployment](#group-policy-deployment)
  - [Remote Management](#remote-management)
- [🏢 Enterprise Usage](#-enterprise-usage)
  - [Deployment Strategies](#deployment-strategies)
  - [Monitoring & Reporting](#monitoring--reporting)
  - [Best Practices](#best-practices)
- [🔍 Troubleshooting](#-troubleshooting)
  - [Common Issues](#common-issues)
  - [Debugging Commands](#debugging-commands)
  - [Recovery Procedures](#recovery-procedures)
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

```batch
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

## ✨ Key Features

### 🛡️ **Enterprise-Grade Safety**
- **Administrative Privilege Verification**: Automatically checks for required admin rights
- **Service State Management**: Proper service stop/start with error handling
- **File System Protection**: Safe deletion of spool files with error recovery
- **Exit Code Standards**: Proper return codes for automation and monitoring

### 🎯 **Intelligent Processing**
- **Service Dependency Handling**: Manages Print Spooler service gracefully
- **File Type Targeting**: Specifically removes `.spl` and `.shd` spool files
- **Error Detection**: Identifies and reports specific failure points
- **Timeout Management**: Built-in delays for service state transitions

### 📊 **Comprehensive Logging**
- **Clear Status Messages**: Real-time feedback on operation progress
- **Error Classification**: Detailed error reporting with specific codes
- **Verbose Mode**: Optional detailed logging for troubleshooting
- **Success Confirmation**: Clear completion status reporting

### 🔄 **Automation Ready**
- **Silent Operation**: Can run without user interaction
- **Task Scheduler Compatible**: Perfect for scheduled maintenance
- **Remote Execution**: Supports remote deployment and execution
- **Batch & PowerShell**: Dual implementation for different environments

## 📋 Installation & Usage

### System Requirements

- **Operating System**: Windows 7/8/10/11, Windows Server 2008 R2+
- **Permissions**: Must run as **Administrator**
- **PowerShell Version**: 5.1 or later (for .ps1 script)
- **Dependencies**: Native Windows utilities only

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

| Script | Execution Method | Best For |
|--------|------------------|----------|
| **Clear-PrintSpoolerQueue.bat** | Right-click → Run as administrator | Quick manual cleanup |
| **Clear-PrintSpoolerQueue.ps1** | PowerShell with `-Verbose` | Detailed logging & automation |

### Usage Examples

#### Basic Execution
```batch
# Batch script - simple double-click or:
Clear-PrintSpoolerQueue.bat

# PowerShell script - from elevated PowerShell:
.\Clear-PrintSpoolerQueue.ps1

# PowerShell with verbose output:
.\Clear-PrintSpoolerQueue.ps1 -Verbose
```

#### Remote Execution
```powershell
# Execute on remote computer
Invoke-Command -ComputerName "PrintServer01" -ScriptBlock {
    & "C:\Scripts\Clear-PrintSpoolerQueue.ps1" -Verbose
}

# Execute via PsExec
psexec \\PrintServer01 -s "C:\Scripts\Clear-PrintSpoolerQueue.bat"
```

#### Scheduled Execution
```batch
# Create daily cleanup task
schtasks /create /tn "Daily Print Cleanup" /tr "C:\Scripts\Clear-PrintSpoolerQueue.bat" /sc daily /st 03:00 /ru SYSTEM

# Create weekly cleanup with PowerShell
schtasks /create /tn "Weekly Print Cleanup Verbose" /tr "powershell.exe -ExecutionPolicy Bypass -File C:\Scripts\Clear-PrintSpoolerQueue.ps1 -Verbose" /sc weekly /d SUN /st 02:00 /ru SYSTEM
```

## 🏗️ Advanced Features

### Script Comparison

| Feature | Batch Script | PowerShell Script |
|---------|-------------|-------------------|
| **Admin Check** | `cacls.exe` method | `WindowsPrincipal` check |
| **Service Control** | `net stop/start` | `Stop-Service/Start-Service` |
| **File Cleanup** | `del /F /Q` | `Remove-Item -Force -Recurse` |
| **Error Handling** | Basic exit codes | Advanced try/catch blocks |
| **Logging** | Console output | Verbose parameter support |
| **Dependencies** | Native Windows only | PowerShell 5.1+ |

### Error Handling & Recovery

The scripts provide **comprehensive error handling**:

```powershell
# PowerShell error handling example
try {
    Stop-Service -Name 'Spooler' -Force -ErrorAction Stop
    Start-Sleep -Seconds 3
    Remove-Item -Path "$env:SystemRoot\System32\spool\PRINTERS\*" -Force -Recurse -ErrorAction SilentlyContinue
    Start-Service -Name 'Spooler' -ErrorAction Stop
    Write-Output "Print Spooler queue has been forcibly cleared and service restarted."
} catch {
    Write-Error "Failed to clean print spooler: $($_.Exception.Message)"
    Exit 1
}
```

**Recovery Features:**
- 🔄 **Automatic Service Recovery**: If stop fails, attempts graceful handling
- 📝 **Detailed Error Messages**: Specific error reporting for troubleshooting
- 🛡️ **Safe Fallback**: Service restart attempted even if file deletion fails
- 🚨 **Exit Code Standards**: Proper return codes (0=success, 1=failure)

### Execution Policy Management

If you encounter execution policy errors:

```powershell
# Check current policy
Get-ExecutionPolicy -Scope CurrentUser

# Set policy for current user
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted

# Bypass policy for single execution
powershell.exe -ExecutionPolicy Bypass -File "Clear-PrintSpoolerQueue.ps1"

# Enterprise policy bypass
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "Clear-PrintSpoolerQueue.ps1" -Verbose
```

## 🔗 DevOps Integration

### Task Scheduler Integration

#### GUI Method
1. Open **Task Scheduler** → **Create Task**
2. **General Tab**: Select "Run with highest privileges"
3. **Triggers Tab**: Set desired schedule
4. **Actions Tab**: Configure program execution
5. **Settings Tab**: Enable task settings as needed

#### Command Line Method
```batch
# Create basic daily task
schtasks /create /tn "Print Spooler Cleanup" /tr "C:\Scripts\Clear-PrintSpoolerQueue.bat" /sc daily /st 03:00 /ru SYSTEM /rl HIGHEST

# Create advanced PowerShell task
schtasks /create /tn "Advanced Print Cleanup" ^
    /tr "powershell.exe -NoProfile -ExecutionPolicy Bypass -File 'C:\Scripts\Clear-PrintSpoolerQueue.ps1' -Verbose" ^
    /sc daily /st 03:00 /ru SYSTEM /rl HIGHEST ^
    /f /rl HIGHEST
```

### Group Policy Deployment

#### Startup Script Deployment
```batch
# Copy to SYSVOL
copy Clear-PrintSpoolerQueue.bat "\\domain.com\SYSVOL\domain.com\scripts\"

# Configure via Group Policy:
# Computer Configuration > Policies > Windows Settings > Scripts > Startup
```

#### Scheduled Task via Group Policy
```xml
<!-- Import into Group Policy Preferences > Control Panel Settings > Scheduled Tasks -->
<ScheduledTask clsid="{CC63F200-7309-4ba0-B154-A71CD118DBCC}">
    <Properties action="C" name="Print Spooler Cleanup" runAs="SYSTEM" logonType="S4U"/>
    <Task version="1.3">
        <Actions>
            <Exec>
                <Command>C:\Scripts\Clear-PrintSpoolerQueue.bat</Command>
            </Exec>
        </Actions>
        <Triggers>
            <CalendarTrigger>
                <StartBoundary>2025-01-01T03:00:00</StartBoundary>
                <ScheduleByDay>
                    <DaysInterval>1</DaysInterval>
                </ScheduleByDay>
            </CalendarTrigger>
        </Triggers>
    </Task>
</ScheduledTask>
```

### Remote Management

```powershell
# Execute on multiple computers
$Computers = @("PC001", "PC002", "PrintServer01")
$ScriptPath = "C:\Scripts\Clear-PrintSpoolerQueue.ps1"

Invoke-Command -ComputerName $Computers -ScriptBlock {
    param($Path)
    & $Path -Verbose
} -ArgumentList $ScriptPath

# Using PowerShell Desired State Configuration (DSC)
Configuration PrintSpoolerMaintenance {
    Script ClearSpoolerQueue {
        GetScript = { @{ Result = "PrintSpoolerCleanup" } }
        TestScript = { $false }  # Always run
        SetScript = { & "C:\Scripts\Clear-PrintSpoolerQueue.ps1" }
    }
}
```

## 🏢 Enterprise Usage

### Deployment Strategies

#### Small Office (1-50 computers)
```batch
# Simple network share deployment
net use Z: \\server\scripts
copy Clear-PrintSpoolerQueue.bat Z:\
# Create scheduled task pointing to Z:\Clear-PrintSpoolerQueue.bat
```

#### Medium Enterprise (50-500 computers)
```powershell
# Group Policy + PowerShell deployment
# 1. Deploy script via Group Policy Preferences
# 2. Create scheduled task via Group Policy
# 3. Monitor via event logs
```

#### Large Enterprise (500+ computers)
```powershell
# SCCM/ConfigMgr deployment
# 1. Package scripts as application
# 2. Deploy to device collections
# 3. Schedule recurring deployment
# 4. Report compliance via SCCM reporting
```

### Monitoring & Reporting

#### Event Log Integration
```powershell
# Add to PowerShell script for logging
Write-EventLog -LogName System -Source "Print Spooler Cleanup" -EventId 1001 -EntryType Information -Message "Spooler cleanup completed successfully"

# Query cleanup events
Get-EventLog -LogName System -Source "Print Spooler Cleanup" -Newest 10
```

#### Performance Monitoring
```powershell
# Monitor print spooler performance
Get-Counter "\Print Queue(*)\Jobs" -Continuous
Get-Counter "\Print Queue(*)\Jobs Spooling" -Continuous

# Check spooler service status across domain
Get-ADComputer -Filter * | ForEach-Object {
    Get-Service -Name Spooler -ComputerName $_.Name -ErrorAction SilentlyContinue
}
```

### Best Practices

#### Security Considerations
- ✅ **Run as SYSTEM**: Use system account for scheduled tasks
- ✅ **Secure Script Storage**: Store scripts in protected locations
- ✅ **Execution Policy**: Configure appropriate PowerShell execution policies
- ✅ **Audit Trail**: Enable logging for compliance requirements

#### Operational Guidelines
- 🕐 **Schedule During Off-Hours**: Run cleanup during low-usage periods
- 📊 **Monitor Success Rates**: Track cleanup success/failure rates
- 🔄 **Regular Reviews**: Periodically review and update scripts
- 📝 **Documentation**: Maintain operational documentation

## 🔍 Troubleshooting

### Common Issues

#### Access Denied Errors
```batch
# Problem: "Access is denied" when stopping service
# Solution: Verify administrative privileges
whoami /priv | findstr SeServiceLogonRight
# Ensure "Run as administrator" is used
```

#### Service Won't Stop
```powershell
# Problem: Print Spooler service won't stop
# Solution: Force stop dependent services first
Get-Service | Where-Object {$_.DependentServices -contains (Get-Service Spooler)} | Stop-Service -Force
Stop-Service Spooler -Force

# Alternative: Use taskkill for spoolsv.exe
taskkill /f /im spoolsv.exe
```

#### PowerShell Execution Policy
```powershell
# Problem: Execution policy blocks script
# Solution: Check and modify execution policy
Get-ExecutionPolicy -List
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Emergency bypass
powershell.exe -ExecutionPolicy Bypass -File "Clear-PrintSpoolerQueue.ps1"
```

### Debugging Commands

```batch
# Check spooler service status
sc query spooler
net start | findstr Spooler

# List files in spool directory
dir "%SystemRoot%\System32\spool\PRINTERS\" /a

# Check print queues
wmic printer get Name,PrintJobCount
```

```powershell
# PowerShell debugging
Get-Service Spooler | Format-List *
Get-PrintJob -PrinterName * | Format-Table
Get-ChildItem "$env:SystemRoot\System32\spool\PRINTERS" | Measure-Object
```

### Recovery Procedures

#### Manual Service Recovery
```batch
# If automatic restart fails
net stop spooler /y
timeout /t 5
net start spooler

# Check service dependencies
sc enumdepend spooler
```

#### Registry Cleanup (Advanced)
```reg
# Clear spooler registry entries (use with caution)
# HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Print\Printers
# Only modify if standard cleanup fails
```

## 🤝 Contributing

We welcome contributions! Here's how to get involved:

### Development Setup
```bash
git clone https://github.com/paulmann/Print-Spooler-Queue-Cleanup.git
cd Print-Spooler-Queue-Cleanup

# Test batch script syntax
# No syntax checker needed - test by execution

# Test PowerShell script syntax
powershell.exe -NoProfile -NoExecute -Command "& {Get-Command .\Clear-PrintSpoolerQueue.ps1}"
```

### Contribution Guidelines

1. **Fork** the repository
2. **Create** a feature branch (`git checkout -b feature/amazing-feature`)
3. **Test** your changes thoroughly on different Windows versions
4. **Commit** your changes (`git commit -m 'Add amazing feature'`)
5. **Push** to the branch (`git push origin feature/amazing-feature`)
6. **Open** a Pull Request

### Code Standards

- ✅ **Windows Compatibility**: Test on Windows 7-11 and Server versions
- ✅ **Error Handling**: Comprehensive error checking and recovery
- ✅ **Documentation**: Comment complex logic and add help text
- ✅ **Backwards Compatibility**: Maintain compatibility with existing deployments
- ✅ **Security**: Follow security best practices for system administration

## 📄 License

This project is licensed under the **MIT License** - see the [LICENSE](LICENSE) file for details.

```
MIT License

Copyright (c) 2025 Mikhail Deynekin

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.
```

## 👨‍💻 Author & Support

**Mikhail Deynekin**  
Senior System Administrator & Developer

- 🌐 Website: [deynekin.com](https://deynekin.com)
- 📧 Email: mid1977@gmail.com
- 🐙 GitHub: [@paulmann](https://github.com/paulmann)

### Getting Help

- 📖 **Documentation**: Read this README thoroughly
- 🐛 **Bug Reports**: [Open an issue](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/issues/new)
- 💡 **Feature Requests**: [Request features](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/issues/new)
- 💬 **Questions**: Check [Discussions](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/discussions)

### Related Projects

- [Print-Management-Tools](https://github.com/search?q=print+spooler+windows) - Various Windows print management utilities
- [Windows-System-Administration](https://github.com/topics/windows-administration) - General Windows admin tools
- [PowerShell-Scripts](https://github.com/topics/powershell-scripts) - Collection of PowerShell utilities

---

### ⭐ Star this repository if it helped you!

**Print Spooler Queue Cleanup** - *Keeping your printers flowing, one cleanup at a time* 🖨️🧹

[Report Bug](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/issues) · [Request Feature](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/issues) · [Documentation](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki)

# Getting Started

Five minutes from download to your first safe cleanup.

---

## 1. Get the files

You need exactly one file to run the tool. The `.bat` is optional and only exists so the tool can
be double-clicked.

```powershell
# From the latest release
Invoke-WebRequest https://github.com/paulmann/Print-Spooler-Queue-Cleanup/releases/latest/download/Clear-PrintSpoolerQueue.ps1 -OutFile Clear-PrintSpoolerQueue.ps1
Invoke-WebRequest https://github.com/paulmann/Print-Spooler-Queue-Cleanup/releases/latest/download/Clear-PrintSpoolerQueue.bat -OutFile Clear-PrintSpoolerQueue.bat

# Windows marks files downloaded from the internet; clear the flag or the
# execution policy will block the script.
Unblock-File .\Clear-PrintSpoolerQueue.ps1
```

Or clone the repository:

```powershell
git clone https://github.com/paulmann/Print-Spooler-Queue-Cleanup.git
cd Print-Spooler-Queue-Cleanup
```

Verify what you downloaded before running it on a print server:

```powershell
Get-FileHash .\Clear-PrintSpoolerQueue.ps1 -Algorithm SHA256
# compare with checksums.txt on the release page
```

**Requirements:** Windows 10/11 or Server 2012 R2+, Windows PowerShell 5.1 or PowerShell 7.x,
and Administrator rights for the local machine. No modules to install.

---

## 2. Look before you touch

The first command you should ever run changes nothing. It does not even stop the service.

```powershell
.\Clear-PrintSpoolerQueue.ps1 -WhatIf
```

You get the full inventory — how many jobs, which printers, who submitted them, how big, how old —
and the exact plan, with no side effects at all.

Then find out *why* it is stuck:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic
```

This is read-only and safe on production at any time. It reports the Spooler and its dependent
services, `spoolsv.exe` responsiveness and memory, every printer with its offline/error state,
installed drivers, TCP/IP ports, the spool directory contents and the age of the oldest file, the
queued-job table, and the relevant event-log entries. Most tickets are explained by this one screen.
See the [Diagnostics Guide](Diagnostics-Guide) for how to read it.

---

## 3. Fix it

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Force -Backup
```

That is the whole operation. What actually happens:

1. The tool resolves the real spool directory (registry override, then OS default) and runs four
   safety guards against it.
2. It inventories the queue from `Win32_PrintJob`, falling back to `.SHD` headers when the Spooler
   is too hung to answer WMI.
3. It archives every file it is about to delete into a ZIP with a JSON manifest and a SHA-256.
4. It stops `PrintWorkflowUserSvc_*`, `Fax` and `PrintNotify` **first**, then the Spooler — which
   is what makes the stop succeed instead of hanging.
5. It deletes `*.SPL`/`*.SHD`, retrying locked files with exponential backoff.
6. It restarts the Spooler, then only the dependent services that were actually running.
7. It re-reads the service state and re-scans the directory to confirm the queue did not refill.
8. It prints a summary and sets a documented exit code.

Without `-Force` you are prompted first. Without `-Backup` nothing is archived — on a print server
with legally or financially significant queues, use `-Backup`, or `-BackupRequired` to make a
failed backup abort the run.

---

## 4. Read the result

```
==============================================================================
    EXECUTION SUMMARY
==============================================================================
    Overall result      : Success
    Mode                : Cleanup
    Duration            : 00:00:14.882
    Targets             : 1 attempted / 1 succeeded / 0 warned / 0 failed
    Print jobs found    : 7
    Jobs removed        : 7
    Files deleted       : 14
    Space reclaimed     : 184.32 MB
    Warnings / errors   : 0 / 0
    Backup(s)           : C:\ProgramData\PrintSpoolerCleanup\Backups\spool-backup_SRV01_20261010-031500.zip
    Log file            : C:\ProgramData\PrintSpoolerCleanup\Logs\PrintSpoolerCleanup.log
    Exit code           : 0
==============================================================================
```

`exit 0` means every target was processed and the Spooler was verified running. The full contract
is in [Exit Codes](Exit-Codes).

---

## The four commands you will actually use

| Situation | Command |
|---|---|
| **What is wrong?** | `.\Clear-PrintSpoolerQueue.ps1 -Diagnostic` |
| **Fix it safely** | `.\Clear-PrintSpoolerQueue.ps1 -Force -Backup` |
| **Fix it without killing healthy jobs** | `.\Clear-PrintSpoolerQueue.ps1 -OlderThan 2h -Force -Backup` |
| **It keeps coming back** | `.\Clear-PrintSpoolerQueue.ps1 -Repair -Force -Backup` |

---

## Working on a busy shared queue?

This is the single most useful filter in the tool. A print server with 40 jobs usually has *one*
wedged job and 39 that are fine. Deleting all 40 solves your ticket and creates thirty new ones.

```powershell
# Only jobs that have been sitting for more than two hours
.\Clear-PrintSpoolerQueue.ps1 -OlderThan 2h -Force -Backup

# Only one printer, everything else untouched
.\Clear-PrintSpoolerQueue.ps1 -PrinterName "KYOCERA TASKalfa*" -Force -Backup

# Only the jobs a specific user submitted
.\Clear-PrintSpoolerQueue.ps1 -OwnerName "CORP\jdoe" -Force -Backup

# Only the specific job ids you saw in -Diagnostic
.\Clear-PrintSpoolerQueue.ps1 -JobId 4217,4218 -Force -Backup
```

Accepts `30s`, `15m`, `2h`, `7d`, `1w`, a `[TimeSpan]`, or `01:30:00`.

> **One caveat worth knowing:** when the Spooler is completely hung it cannot answer WMI, so the
> inventory is built from `.SHD` headers, where printer and document names are best-effort. Jobs
> with no readable metadata are **not** selected by name-based filters, and the report says so in
> its warnings. That is deliberate: silently guessing which printer a job belongs to and then
> deleting it would be worse than doing nothing.

---

## Double-click users

Run `Clear-PrintSpoolerQueue.bat`. It:

- prefers PowerShell 7 (`pwsh.exe`) and falls back to Windows PowerShell 5.1
- can request UAC elevation itself: `Clear-PrintSpoolerQueue.bat /Elevate`
- applies any staged update before running
- **never** pauses when launched by Task Scheduler
- uses exactly the same exit codes as the `.ps1`

```cmd
Clear-PrintSpoolerQueue.bat /?                     rem help
Clear-PrintSpoolerQueue.bat /Diagnostic            rem read-only
Clear-PrintSpoolerQueue.bat /Force /Backup         rem fix it
Clear-PrintSpoolerQueue.bat /Force /Backup /OlderThan 2h
```

Full switch list: [Batch Launcher Reference](Batch-Launcher-Reference).

---

## The complete manual, in your console

The built-in help is generated from the live parameter metadata, so it cannot drift away from the
code the way a hand-written README does.

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Help                    # everything (~60 KB)
.\Clear-PrintSpoolerQueue.ps1 -ListHelpSections       # what sections exist
.\Clear-PrintSpoolerQueue.ps1 -HelpSection Safety      # just the safety model
.\Clear-PrintSpoolerQueue.ps1 -HelpSection Troubleshooting  # symptom -> cause -> fix
Get-SpoolerCleanupHelp -Search backup                  # every block mentioning backups
```

Sections: `Usage`, `QuickStart`, `Modes`, `Parameters`, `Examples`, `ExitCodes`, `Configuration`,
`Environment`, `Logging`, `Reports`, `Remote`, `Scheduling`, `Updates`, `Safety`,
`Troubleshooting`, `Diagnostics`, `Version`, `Credits`.

---

## Next steps

| You want to… | Go to |
|---|---|
| Understand what a stuck queue actually is | [Print Spooler Internals](Print-Spooler-Internals) |
| Run it against 400 machines | [Remote and Fleet Management](Remote-and-Fleet-Management) |
| Make it run every night | [Scheduling and Automation](Scheduling-and-Automation) |
| Feed results into your RMM | [Reports and Machine-Readable Output](Reports-and-Machine-Readable-Output) |
| Set site-wide defaults | [Configuration](Configuration) |
| Know exactly what it will and will not delete | [Safety Model](Safety-Model) |
| Keep the tool current | [Updates and Versioning](Updates-and-Versioning) |
| Fix something that went wrong | [Troubleshooting](Troubleshooting) |

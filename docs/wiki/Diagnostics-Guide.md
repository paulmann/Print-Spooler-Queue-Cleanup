# Diagnostics Guide

`-Diagnostic` is read-only. It never stops a service and never touches a file, so it is safe on a
production print server at any time — including during the incident you are currently in.

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic -ExportFormat Html -ExportPath C:\Temp
.\Clear-PrintSpoolerQueue.ps1 -ComputerName SRV-PRINT01, SRV-PRINT02 -Diagnostic -ExportFormat Csv
Get-PrintSpoolerDiagnostic                                   # the cmdlet form
```

Exit code **11** means "diagnostics completed". It is informational, not an error — do not alert
on it.

---

## What it collects

| Area | Fields |
|---|---|
| **Operating system** | Caption, version, build, last boot time, free memory |
| **Services** | `Spooler`, `PrintNotify`, `Fax`, `PrintWorkflowUserSvc_*` — state, start mode, process id |
| **Processes** | `spoolsv.exe` and `splwow64.exe` — CPU seconds, working set, thread count, responding flag, start time |
| **Printers** | Name, state, status, local/network/shared, **offline flag**, port, driver |
| **Drivers** | Name, manufacturer, version, INF |
| **Ports** | TCP/IP printer ports — host address, port number, protocol, SNMP |
| **Spool directory** | Path, file count, total bytes, per-extension breakdown, oldest/newest file age, ACL owner |
| **Queued jobs** | Id, printer, document, owner, size, pages, status, submission time, **data source** |
| **Registry** | `DefaultSpoolDirectory`, print event logging, driver isolation settings |
| **Event logs** | Service Control Manager entries mentioning Spooler/Print, plus `PrintService/Admin` |

Control the event-log window with `-EventLogHours` (default 72) and
`Diagnostics.EventLogMaxEntries` (default 200).

---

## The six patterns that explain most tickets

### Pattern 1 — `Spooler = Running`, queue depth > 0, oldest job very old

The classic wedge. A job that cannot be rendered is blocking everything behind it.

```
Oldest file    : F00412.SHD (37.5 h old)
Queued jobs    : 14
```

**Action.** Clear with a filter if the queue is shared:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -OlderThan 2h -Force -Backup
```

See [Why Print Queues Get Stuck → cause 1](Why-Print-Queues-Get-Stuck#1-a-job-that-cannot-be-rendered).

---

### Pattern 2 — `Spooler = Stop Pending`

The stop is hung. Look at `PrintNotify` and at `splwow64.exe` in the same report: if `PrintNotify`
is `Running`, or `splwow64.exe` shows `Responding = False` with a high thread count, you have found
it.

**Action.**

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Repair -Force -Backup
```

Dependents go down first, then the Spooler, with a process-kill fallback if the stop still will not
complete.

---

### Pattern 3 — Every job's `Source` is `SHD`

The Spooler is too sick to answer WMI, so the inventory was built from `.SHD` headers instead.
Metadata (printer and document names) is best-effort in this mode.

**Consequence you must know:** name-based filters (`-PrinterName`, `-DocumentName`, `-OwnerName`)
will **not** match jobs with no readable metadata, and those jobs are deliberately not selected.
The report warns about it rather than guessing.

**Action.** Either filter by `-JobId` (always available, derived from the file name) or by
`-OlderThan` (uses file timestamps), or drop the filter entirely when the whole queue must go.

---

### Pattern 4 — A printer with `Offline = YES`

Not a queue problem. The queue is paused, or the client has "Use Printer Offline" set, or the device
stopped responding.

**Action.** Clearing the queue will look like it worked for about an hour. Fix the printer state
first:

```powershell
Get-Printer -Name "<printer>" | Select-Object Name, PrinterStatus, WorkOffline, PortName
Set-Printer -Name "<printer>" -PrinterStatus Normal
```

---

### Pattern 5 — Repeated SCM 7031 / 7034 events

The Spooler crashed and was restarted — possibly many times. Usually a driver.

```
Recent Service Control Manager events
  2026-10-09 03:14:02  [Error] id=7031  The Print Spooler service terminated unexpectedly...
  2026-10-09 03:14:02  [Error] id=7031  ...it has done this 4 time(s)...
```

**Action.** Deep clean, then schedule self-healing so a crash stops generating tickets:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Repair -PurgeFormsCache -PurgeDriverStore -Force -Backup

Install-SpoolerCleanupTask -ScriptPath C:\Tools\Clear-PrintSpoolerQueue.ps1 -Schedule OnEvent `
    -EventLogName System -EventSource "Service Control Manager" -EventId 7031
```

---

### Pattern 6 — `SpoolDirectory.Path` is not what you expected

The `DefaultSpoolDirectory` registry override has been changed, or is corrupt.

```
Spool directory : \\filesrv\spool\PRINTERS (Registry)
```

**Action.** Inspect and reset:

```powershell
Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers' -Name DefaultSpoolDirectory
.\Clear-PrintSpoolerQueue.ps1 -Repair -ResetSpoolDirectory -Force
```

Note that a UNC spool directory is **refused by the guards** — remote file systems must be cleaned
on their owner. That refusal is deliberate: a tool with administrator rights and a recursive delete
pointed at a network share is not something anyone should want.

---

## Reading the job table

```
JobId  Printer              Document                 Owner          Size       Pages Status      Submitted          Source
-----  -------              --------                 -----          ----       ----- ------      ---------          ------
412    KYOCERA TASKalfa 4054 invoice-9931.pdf         CORP\jdoe      4.21 MB    3     Printing    2026-10-08 14:02   WMI+SHD
413    KYOCERA TASKalfa 4054 report-Q3.xlsx           CORP\asmith    18.90 MB   240   Spooled     2026-10-09 09:15   WMI+SHD
414    HP LaserJet M507     <unknown>                               1.02 MB    0     Unknown     2026-10-09 22:40   SHD
```

| Column | How to read it |
|---|---|
| `JobId` | Derived from the `.SHD` file name when WMI is unavailable, so always present |
| `Size` | `.SHD` + `.SPL` combined. A multi-GB job is usually the problem, not the victim |
| `Pages` | 0 means unknown — normal for `SHD`-sourced jobs |
| `Status` | `Spooled` = waiting, `Printing` = active, `Error` = failed, `Unknown` = metadata unavailable |
| `Submitted` | The age filter works on this (or on the file timestamp for `SHD`-sourced jobs) |
| `Source` | **How much to trust the row:** `WMI` authoritative · `WMI+SHD` correlated, best of both · `SHD` best-effort |

A row with `Source = SHD` and a blank printer/document is an honest "we could not read the metadata",
not a bug. The alternative — guessing — would mean deleting jobs the operator did not ask to delete.

---

## Exporting for a ticket

The HTML report is self-contained: inline CSS, no CDN, no JavaScript. It renders identically in a
browser, in Outlook and in a ticketing system, and it cannot execute anything.

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic -ExportFormat Html -ExportPath C:\Temp
```

For a bundle that also carries the environment, effective configuration (secrets redacted), recent
log buffer and file integrity hashes:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic -DiagnosticsBundlePath C:\Temp
```

Attach that ZIP to an
[issue](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/issues) and a fix usually takes one
round trip instead of five.

---

## Diagnostics across a fleet

```powershell
$reports = Get-PrintSpoolerDiagnostic -ComputerName (Get-Content .\hosts.txt) `
                -ExportFormat Csv -ExportPath \\filesrv\logs\spooler

$reports.Results | ForEach-Object {
    [pscustomobject]@{
        Computer = $_.ComputerName
        Jobs     = $_.JobsFound
        OldestH  = ($_.Jobs | Measure-Object -Property SubmittedUtc -Maximum).Maximum
        Bytes    = $_.BytesFound
        Service  = $_.ServiceIsRunning
    }
} | Sort-Object Bytes -Descending | Format-Table
```

For a recurring health signal rather than a deep report, `Test-PrintSpoolerHealth` is cheaper and
returns a 0–100 score:

```powershell
Get-Content .\hosts.txt | Test-PrintSpoolerHealth -ThrottleLimit 16 |
    Where-Object Health -ne 'Healthy' |
    Format-Table ComputerName, Health, Score, QueueDepth, OldestJobHours, Findings -Wrap
```

---

## Related pages

- [Why Print Queues Get Stuck](Why-Print-Queues-Get-Stuck) — the seven root causes
- [Print Spooler Internals](Print-Spooler-Internals) — the mechanisms
- [Repair and Deep Cleaning](Repair-and-Deep-Cleaning)
- [Reports and Machine-Readable Output](Reports-and-Machine-Readable-Output)

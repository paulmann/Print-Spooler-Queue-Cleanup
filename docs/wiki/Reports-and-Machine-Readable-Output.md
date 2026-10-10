# Reports and Machine-Readable Output

Two audiences, two contracts. A human reads the console summary or an HTML report; a machine reads
one JSON document on stdout and a process exit code. Both are stable and versioned.

---

## Exporting files

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Force -Backup -ExportFormat Json,Csv,Html,Text -ExportPath C:\Reports
```

| Format | File | Best for |
|---|---|---|
| `Json` | `.json` | Full fidelity — every field, every job, every warning |
| `Csv` | `.csv` | One row per host: spreadsheets, Power BI, SIEM ingestion |
| `Html` | `.html` | Attaching to a ticket, emailing a manager |
| `Text` | `.txt` | Pasting into a chat, `more`, a plain-text ticket field |

`-ExportPath` may be a **directory** (files are auto-named
`Clear-PrintSpoolerQueue_<stamp>.<ext>`) or a **full file path** (used for the first format; later
formats get suffixed names alongside it).

Defaults can be set site-wide with `Execution.ExportFormat` and `Execution.ExportPath`, so a
scheduled task writes a CSV to your log share on every run without anyone typing it again.

---

## The HTML report

**Self-contained**: inline CSS, no CDN, no external stylesheet, no JavaScript. It renders identically
in a browser, in Outlook's reading pane and in a ticketing system, it works on an air-gapped network,
and it cannot execute anything.

Sections: run header with status badge · run totals · per-computer table · queued jobs discovered ·
warnings and errors · backups created with their SHA-256 · footer with the generating version.

**Every embedded value is HTML-escaped.** A computer named `<script>alert(1)</script>` or a job whose
document name contains markup renders as inert text. This is asserted by a unit test, because a report
that gets pasted into a web-based ticketing system is a stored-XSS vector if it is not.

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic -ExportFormat Html -ExportPath C:\Temp
Invoke-Item C:\Temp\Clear-PrintSpoolerQueue_20261010-031500.html
```

---

## The CSV report

One row per target host — the shape you want for a fleet spreadsheet:

```
RunStartedUtc, ToolVersion, ComputerName, IsLocal, Status, SpoolDirectory,
ServiceWasRunning, ServiceIsRunning, JobsFound, JobsSelected, JobsRemoved,
FilesFound, FilesDeleted, FilesFailed, FilesLocked, BytesFound, BytesDeleted,
FilesAfter, ResurrectedFiles, ServiceStopMs, ServiceStartMs, DurationMs,
BackupPath, BackupHash, WarningCount, ErrorCount, FirstError
```

```powershell
Import-Csv .\Clear-PrintSpoolerQueue_20261010-031500.csv |
    Where-Object Status -ne 'Success' |
    Sort-Object BytesDeleted -Descending |
    Format-Table ComputerName, Status, JobsRemoved, FilesLocked, FirstError
```

A run with zero targets produces an empty string rather than a malformed file.

---

## Machine-readable stdout

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Force -Backup -MachineReadableJson
```

Guarantee: **exactly one JSON document on stdout, and nothing else.** No banner, no progress records,
no colour escape sequences, no log lines. All human output is redirected to the log file.

This is what makes the tool usable from an RMM agent, a CI step or a monitoring plugin that captures
stdout and parses it.

```powershell
$result = .\Clear-PrintSpoolerQueue.ps1 -Force -Backup -OlderThan 6h `
               -NonInteractive -Quiet -NoProgress -MachineReadableJson | ConvertFrom-Json
$result.summary.overallStatus
$result.summary.targetsFailed
$result.summary.results | Where-Object status -ne 'Success'
```

Do **not** combine `-MachineReadableJson` with `-ExportFormat` — that combination is rejected with
exit code 3, because both want to own the output. Use `-ExportPath` for files and keep stdout for the
document.

---

## The JSON schema

`schemaVersion` is `"4.0"`. A monitoring integration should check it, so a contract change is
detected rather than silently mis-parsed.

**Every key is `camelCase` and matching is case-SENSITIVE.** This matters more than it looks.
PowerShell property access is case-insensitive, so `$report.summary.overallStatus` works in PowerShell
whether the wire key is `overallStatus` or `OverallStatus` — which means a PowerShell-only test cannot
detect a casing regression. `jq`, Python and every RMM parser read the raw text, and for them the
casing *is* the contract:

```bash
# works
... -MachineReadableJson | jq -r '.summary.overallStatus'
# returns null - note the capital O
... -MachineReadableJson | jq -r '.summary.OverallStatus'
```

Keys are produced by one mechanical rule (`ConvertTo-CamelCaseName`): lowercase the first letter, and
for a leading run of *N* capitals followed by a lowercase letter, lowercase the first *N-1*. Hence
`PSEdition` → `psEdition` and `PowerShellVersion` → `powerShellVersion`. There is deliberately no
per-field exception list — a single rule is what stops the document and the emitted bytes from
drifting apart field by field.

**Values are never renamed**, only keys. `status` still carries the literal string
`"SuccessWithWarnings"`, and `jobs[].source` is still `"WMI"`, `"WMI+SHD"` or `"SHD"`.

A contract suite in `tests/Unit.Utilities.ps1` compares the emitted wire text against
[Report-Schema.md](../Report-Schema.md) in both directions, case-sensitively, so neither side can
drift without failing the build.

```jsonc
{
  "schemaVersion": "4.0",
  "summary": {
    "tool": "Clear-PrintSpoolerQueue",
    "toolVersion": "4.0.0",
    "startedUtc": "2026-10-10T03:15:00.1234567Z",
    "finishedUtc": "2026-10-10T03:15:14.8820041Z",
    "durationMs": 14758,
    "duration": "00:00:14.758",
    "overallStatus": "PartialSuccess",
    "mode": "Cleanup",
    "dryRun": false,
    "operator": "CORP\\svc-print",
    "hostComputer": "SRV-MGMT01",
    "powerShellVersion": "7.6.6",
    "psEdition": "Core",
    "logPath": "C:\\ProgramData\\PrintSpoolerCleanup\\Logs\\PrintSpoolerCleanup.log",
    "configSource": "AutoDiscover",

    "targetsAttempted": 12, "targetsSucceeded": 10, "targetsWithWarnings": 1,
    "targetsNothingToDo": 0, "targetsDryRun": 0, "targetsDiagnostic": 0,
    "targetsFailed": 1, "targetsSkipped": 0,

    "jobsFound": 34, "jobsSelected": 34, "jobsRemoved": 33,
    "filesFound": 68, "filesDeleted": 66, "filesFailed": 2, "filesLocked": 2,
    "resurrectedFiles": 0,
    "bytesFound": 433127424, "bytesDeleted": 412876800,

    "warnings": 3, "errors": 1,
    "backupsCreated": ["C:\\ProgramData\\...\\spool-backup_SRV01_....zip"],
    "exitCode": 1,

    "results": [ /* one record per computer, see below */ ]
  }
}
```

### Per-computer result

| Field | Type | Notes |
|---|---|---|
| `computerName` | string | |
| `status` | string | `Success`, `SuccessWithWarnings`, `NothingToDo`, `DryRun`, `Diagnostic`, `Failed`, `ServiceFailure`, `Cancelled` |
| `startedUtc`, `finishedUtc` | ISO-8601 | |
| `durationMs` | long | Time on the target |
| `isLocal` | bool | |
| `spoolDirectory` | string | Resolved path |
| `spoolDirectorySource` | string | `Registry`, `OSDefault`, `Parameter` |
| `serviceWasRunning`, `serviceIsRunning` | bool? | `null` when it could not be determined |
| `serviceStopMs`, `serviceStartMs` | long | Useful for spotting a slow stop |
| `dependentServices` | string[] | Which dependents were cycled |
| `jobsFound`, `jobsSelected`, `jobsRemoved` | int | |
| `filesFound`, `filesSelected`, `filesDeleted`, `filesFailed`, `filesLocked` | int | |
| `bytesFound`, `bytesDeleted` | long | |
| `filesAfter` | int | Spool files remaining after restart |
| `resurrectedFiles` | int | Files that reappeared — a driver is resubmitting |
| `backupPath`, `backupHash` | string? | SHA-256 of the archive |
| `dryRun`, `repairPerformed`, `diagnosticsCollected` | bool | |
| `warnings[]`, `errors[]` | string[] | Verbatim, actionable |
| `jobs[]` | object[] | `jobId`, `printerName`, `documentName`, `userName`, `sizeBytes`, `totalPages`, `pagesPrinted`, `status`, `priority`, `submittedUtc`, `source`, `shdFile`, `splFile` |
| `removedFiles[]` | object[] | `file`, `bytes`, `attempts` |
| `failedFiles[]` | object[] | `file`, `bytes`, `locked`, `attempts`, `error` |
| `orphanedFiles[]` | string[] | `.SHD`/`.SPL` with no matching WMI job |
| `serviceTrace[]` | object[] | `service`, `action`, `from`, `to`, `success`, `timeUtc`, `note` |
| `diagnostics` | object? | Present in `-Diagnostic` mode only |

`jobs[].source` is `WMI`, `WMI+SHD` or `SHD` — it tells you how much to trust that row's metadata.
See [Diagnostics Guide → Reading the job table](Diagnostics-Guide#reading-the-job-table).

---

## `-PassThru`: the same data, in-process

```powershell
$run = .\Clear-PrintSpoolerQueue.ps1 -ComputerName $hosts -Force -Backup -PassThru

$run.OverallStatus
$run.ExitCode
$run.Results | Where-Object { $_.ServiceIsRunning -eq $false }
$run.Results | ForEach-Object { $_.ServiceTrace } | Where-Object { -not $_.Success }
```

`-PassThru` returns the identical object graph the JSON carries, without a serialise/parse round trip.
Prefer it in PowerShell automation; prefer `-MachineReadableJson` when the consumer is another
process.

---

## Monitoring integration patterns

### Nagios / Zabbix / check_mk plugin

```powershell
$json = & .\Clear-PrintSpoolerQueue.ps1 -Force -Backup -OlderThan 6h `
            -NonInteractive -Quiet -MachineReadableJson | ConvertFrom-Json
Write-Output ("{0} | jobs={1} files={2} locked={3} bytes={4}" -f `
    $json.summary.overallStatus, $json.summary.jobsRemoved, $json.summary.filesDeleted,
    $json.summary.filesLocked, $json.summary.bytesDeleted)
exit $json.summary.exitCode
```

The exit code already implements the standard 0/1/2 severity convention closely enough to be used
directly — see [Exit Codes](Exit-Codes).

### Health polling without cleaning

Cheaper than a full cleanup, and non-destructive:

```powershell
Get-Content .\hosts.txt | Test-PrintSpoolerHealth -ThrottleLimit 16 |
    Select-Object ComputerName, Health, Score, QueueDepth, OldestJobHours |
    ConvertTo-Json
```

### Long-term trending

```powershell
.\Clear-PrintSpoolerQueue.ps1 -ComputerName $hosts -Force -Backup `
    -ExportFormat Csv -ExportPath \\filesrv\logs\spooler
```

Append the CSVs into a data store and you get "which print servers wedge most often", "which drivers
correlate with locked files", and "how much spool volume does the site generate" — all from data you
were already producing.

---

## The text report

For pasting into a chat or a ticket field:

```
==============================================================================
Clear-PrintSpoolerQueue v4.0.0 - run report
==============================================================================
Started (UTC) : 2026-10-10T03:15:00.1234567Z
Duration      : 00:00:14.758
Overall       : PartialSuccess
Targets       : 12 attempted, 10 succeeded, 1 warned, 1 failed
Jobs          : 34 found, 34 selected, 33 removed
Bytes         : 413.06 MB found, 393.75 MB reclaimed
Exit code     : 1

------------------------------------------------------------------------------
SRV-PRINT01  [Success]
    Spool directory : C:\Windows\System32\spool\PRINTERS (Registry)
    Service         : was Running, now Running
    Jobs            : found 7, selected 7, removed 7
    Files           : deleted 14, failed 0, locked 0, remaining 0
    Bytes reclaimed : 184.32 MB
    Backup          : C:\ProgramData\...\spool-backup_SRV-PRINT01_....zip (SHA-256 9f2c1a...)
    Dependents cycled: PrintWorkflowUserSvc_1a2b3c, PrintNotify
```

---

## Related pages

- [Exit Codes](Exit-Codes)
- [Logging and Event Log](Logging-and-Event-Log)
- [Diagnostics Guide](Diagnostics-Guide)
- [Remote and Fleet Management](Remote-and-Fleet-Management)

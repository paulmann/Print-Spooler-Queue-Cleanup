# Report Schema

**Status:** normative. `schemaVersion` is `"4.0"`. Consumers should check it, so a contract change is
detected rather than silently mis-parsed.

This document defines the JSON produced by `-MachineReadableJson`, `-ExportFormat Json`, and the
object returned by `-PassThru`. The CSV, HTML and Text reports are projections of the same graph.

---

## Envelope

```jsonc
{
  "schemaVersion": "4.0",
  "summary": { /* run summary, see below */ }
}
```

Produced by `ConvertTo-JsonReport` with `-Depth 8`. Serialised with `ConvertTo-Json`, so on Windows
PowerShell 5.1 non-ASCII characters appear as `\uXXXX` escapes — parse, do not string-match.

**Every key in the document is `camelCase`, and matching is case-SENSITIVE.** PowerShell property
access is case-insensitive, so `$r.summary.OverallStatus` works in PowerShell even when the wire key
is `overallStatus` — which is exactly why this contract cannot be verified from PowerShell alone.
`jq`, Python and every RMM parser see the raw text, and for them the casing is the contract. The
conversion is one mechanical rule (`ConvertTo-CamelCaseName`): lowercase the first letter, and for a
leading run of N capitals followed by a lowercase letter, lowercase the first N-1 — so `PSEdition`
becomes `psEdition` and `PowerShellVersion` becomes `powerShellVersion`. **Values are never renamed**,
only keys: `status` still holds the string `"SuccessWithWarnings"`.

---

## Run summary

| Field | Type | Notes |
|---|---|---|
| `tool` | string | Always `Clear-PrintSpoolerQueue` |
| `toolVersion` | string | SemVer of the running copy |
| `startedUtc` | ISO-8601 | Run start |
| `finishedUtc` | ISO-8601 | Run end |
| `durationMs` | long | |
| `duration` | string | Human form, e.g. `00:00:14.758` |
| `overallStatus` | string | See the status table below |
| `mode` | string | `Cleanup`, `Diagnostic`, `Repair`, `Restore`, `SelfTest` |
| `dryRun` | bool | True under `-WhatIf` |
| `operator` | string | Security principal that ran the tool |
| `hostComputer` | string | Machine the tool ran **from** (not the targets) |
| `powerShellVersion` | string | |
| `psEdition` | string | `Core` or `Desktop` |
| `logPath` | string? | Resolved log file; null when logging was unavailable |
| `configSource` | string | `Defaults`, `Parameter`, `AutoDiscover`, `Environment`, or a `+`-joined combination |
| `targetsAttempted` | int | |
| `targetsSucceeded` | int | |
| `targetsWithWarnings` | int | |
| `targetsNothingToDo` | int | |
| `targetsDryRun` | int | |
| `targetsDiagnostic` | int | |
| `targetsFailed` | int | |
| `targetsSkipped` | int | |
| `jobsFound` | int | Summed across targets |
| `jobsSelected` | int | |
| `jobsRemoved` | int | |
| `filesFound` | int | |
| `filesDeleted` | int | |
| `filesFailed` | int | |
| `filesLocked` | int | |
| `resurrectedFiles` | int | Files that reappeared after restart |
| `bytesFound` | long | |
| `bytesDeleted` | long | |
| `warnings` | int | Count across all targets |
| `errors` | int | Count across all targets |
| `backupsCreated` | string[] | Absolute paths of archives created |
| `exitCode` | int | See [Exit Codes](./Exit-Codes.md) |
| `results` | object[] | One record per target |

### `overallStatus` values

`Success` · `SuccessWithWarnings` · `PartialSuccess` · `Failed` · `Cancelled` · `DryRun` ·
`Diagnostic` · `NothingToDo`

---

## Per-computer result

| Field | Type | Notes |
|---|---|---|
| `computerName` | string | As supplied, or the remote host's own name |
| `status` | string | `Success`, `SuccessWithWarnings`, `NothingToDo`, `DryRun`, `Diagnostic`, `Failed`, `ServiceFailure`, `Cancelled` |
| `startedUtc` / `finishedUtc` | ISO-8601 | Measured **on the target** |
| `durationMs` | long | Time on the target, excluding dispatch overhead |
| `isLocal` | bool | |
| `spoolDirectory` | string | Resolved path actually operated on |
| `spoolDirectorySource` | string | `Registry`, `OSDefault`, `Parameter` |
| `serviceWasRunning` | bool? | `null` when it could not be determined |
| `serviceIsRunning` | bool? | Post-run state. `false` here is the urgent signal |
| `serviceStopMs` | long | |
| `serviceStartMs` | long | |
| `dependentServices` | string[] | Dependents that were cycled (only those that were running) |
| `jobsFound` | int | |
| `jobsSelected` | int | After filters |
| `jobsRemoved` | int | |
| `filesFound` | int | All `.SPL`/`.SHD` present |
| `filesSelected` | int | |
| `filesDeleted` | int | |
| `filesFailed` | int | |
| `filesLocked` | int | Subset of `filesFailed` classified as a sharing violation |
| `bytesFound` | long | |
| `bytesDeleted` | long | |
| `filesAfter` | int | Spool files remaining after restart |
| `resurrectedFiles` | int | Non-zero means a driver or client is resubmitting |
| `backupPath` | string? | |
| `backupHash` | string? | SHA-256 of the archive |
| `dryRun` | bool | |
| `repairPerformed` | bool | |
| `diagnosticsCollected` | bool | |
| `mode` | string | |
| `warnings[]` | string[] | Verbatim, actionable text |
| `errors[]` | string[] | Verbatim, actionable text |
| `jobs[]` | object[] | See below |
| `removedFiles[]` | object[] | `file`, `bytes`, `attempts` |
| `failedFiles[]` | object[] | `file`, `bytes`, `locked`, `attempts`, `error` |
| `orphanedFiles[]` | string[] | `.SHD`/`.SPL` with no matching WMI job |
| `serviceTrace[]` | object[] | `service`, `action`, `from`, `to`, `success`, `timeUtc`, `note` |
| `diagnostics` | object? | `-Diagnostic` mode only |
| `selfTestPrint` | object? | Present only with `-SelfTestPrint` |

### `jobs[]`

| Field | Type | Notes |
|---|---|---|
| `jobId` | long | From WMI, or parsed from the `.SHD` file name |
| `printerName` | string? | Best-effort when `source` is `SHD` |
| `documentName` | string? | |
| `userName` | string? | Submitting user |
| `sizeBytes` | long | `.SHD` + `.SPL` combined |
| `totalPages` | int | 0 when unknown |
| `pagesPrinted` | int | |
| `status` | string | `Spooled`, `Printing`, `Error`, `Unknown` |
| `priority` | int | |
| `submittedUtc` | ISO-8601? | Falls back to the `.SHD` last-write time |
| `source` | string | **How much to trust the row** — see below |
| `shdFile` | string? | |
| `splFile` | string? | |

`source` values:

| Value | Meaning |
|---|---|
| `WMI` | Authoritative, from `Win32_PrintJob` |
| `WMI+SHD` | Correlated: WMI metadata with exact file names from the `.SHD` scan |
| `SHD` | Best-effort, extracted from the `.SHD` header because the Spooler could not answer WMI |

Name-based filters (`-PrinterName`, `-DocumentName`, `-OwnerName`) are **not** applied to jobs whose
metadata could not be read; those jobs are left alone and the fact is recorded in `warnings[]`. A
consumer should surface that warning rather than treating a low `jobsSelected` as a bug.

### `diagnostics` (when `-Diagnostic`)

A dictionary with the keys: `os`, `services`, `spoolerProcess`, `printers`, `drivers`, `ports`,
`spoolDirectory`, `registry`, `eventLog`, `printEventLog`. Shapes are documented in the
[Diagnostics Guide](../wiki/Diagnostics-Guide.md#what-it-collects).

`spoolDirectory.byExtension` is an array of `{ ext, count, bytes }`. `oldestFile` and `newestFile`
are `{ name, ageHours }`.

These are camelCase like every other key in the document, because the whole graph is projected
through one mechanical rule (`ConvertTo-CamelCaseDeep`) before serialisation. There is deliberately
no per-field exception list: a single rule is what keeps this document and the emitted JSON from
drifting apart field by field.

---

## The machine-readable guarantee

`-MachineReadableJson` promises **exactly one JSON document on stdout and nothing else**:

- the banner is suppressed
- progress records are suppressed
- console colours and glyphs are disabled
- log records are not written to the console
- warnings that would have gone to the host go to the log file instead

All human-readable output is redirected to the log file, so an agent can capture stdout and parse it
without stripping decoration.

`-MachineReadableJson` combined with `-ExportFormat` is **rejected with exit code 3**, because both
want to own the output. Use `-ExportPath` for files and keep stdout for the document.

---

## CSV projection

One row per target host. Columns, in order:

```
RunStartedUtc, ToolVersion, ComputerName, IsLocal, Status, SpoolDirectory,
ServiceWasRunning, ServiceIsRunning, JobsFound, JobsSelected, JobsRemoved,
FilesFound, FilesDeleted, FilesFailed, FilesLocked, BytesFound, BytesDeleted,
FilesAfter, ResurrectedFiles, ServiceStopMs, ServiceStartMs, DurationMs,
BackupPath, BackupHash, WarningCount, ErrorCount, FirstError
```

Produced with `ConvertTo-Csv -NoTypeInformation`, so values are quoted only when necessary and
embedded commas/newlines are handled correctly. A run with zero targets yields an empty string rather
than a malformed file.

`FirstError` is the first entry of `errors[]`, which is the field most worth putting in a spreadsheet
column.

---

## HTML projection

Self-contained: inline CSS, **no** CDN, no external stylesheet, **no JavaScript**. It renders
identically in a browser, in Outlook's reading pane and in a ticketing system, works on an air-gapped
network, and cannot execute anything.

**Every embedded value is HTML-escaped** via `System.Net.WebUtility.HtmlEncode`. A computer named
`<script>alert(1)</script>` or a document name containing markup renders as inert text. This is
asserted by a unit test, because a report that gets pasted into a web-based ticketing system is a
stored-XSS vector if it is not.

Sections: header with status badge · run totals · per-computer table · queued jobs discovered ·
warnings and errors · backups created with SHA-256 · footer with the generating version.

---

## Text projection

Fixed-width plain text for pasting into a chat or a ticket field. Truncates the job list at 50 rows
per host; use JSON for full fidelity.

---

## Consumption examples

### PowerShell

```powershell
$result = .\Clear-PrintSpoolerQueue.ps1 -Force -Backup -MachineReadableJson | ConvertFrom-Json
$result.summary.overallStatus
$result.summary.results | Where-Object status -ne 'Success'
$result.summary.results | Where-Object { $_.serviceIsRunning -eq $false }
```

> **Note for PowerShell 7 consumers:** `ConvertFrom-Json` automatically converts ISO-8601 strings
> into `[datetime]`. Asserting on the *wire format* requires the raw text, not the parsed property.

### jq

```bash
.\Clear-PrintSpoolerQueue.ps1 -Force -Backup -MachineReadableJson \
  | jq -r '.summary.results[] | select(.status != "Success") | "\(.computerName)\t\(.status)\t\(.errors[0])"'
```

### Python

```python
import json, subprocess
out = subprocess.run(
    ["pwsh", "-NoProfile", "-File", "Clear-PrintSpoolerQueue.ps1",
     "-Force", "-Backup", "-MachineReadableJson"],
    capture_output=True, text=True)
report = json.loads(out.stdout)
print(report["summary"]["overallStatus"], out.returncode)
for r in report["summary"]["results"]:
    if r["status"] != "Success":
        print(r["computerName"], r["errors"])
```

### Nagios / Zabbix plugin

```powershell
$json = & .\Clear-PrintSpoolerQueue.ps1 -Force -Backup -OlderThan 6h `
            -NonInteractive -Quiet -MachineReadableJson | ConvertFrom-Json
Write-Output ("{0} | jobs={1} files={2} locked={3} bytes={4}" -f `
    $json.summary.overallStatus, $json.summary.jobsRemoved, $json.summary.filesDeleted,
    $json.summary.filesLocked, $json.summary.bytesDeleted)
exit $json.summary.exitCode
```

---

## Versioning rules

| Change | Requires |
|---|---|
| Adding a field | Minor `schemaVersion` note in the changelog; consumers must tolerate unknown fields |
| Removing or renaming a field | **Major** `schemaVersion` bump and a major product version |
| Changing a field's type or unit | **Major** `schemaVersion` bump |
| Changing an `overallStatus` or `status` value's spelling | **Major** `schemaVersion` bump |
| Changing an exit code's meaning | **Major** product version (see [Exit Codes](./Exit-Codes.md)) |

Consumers should ignore unknown fields rather than failing on them, which keeps additive changes
non-breaking in both directions.

---

## See also

- [Exit Codes](./Exit-Codes.md)
- [Wiki: Reports and Machine-Readable Output](../wiki/Reports-and-Machine-Readable-Output.md)
- [Wiki: Diagnostics Guide](../wiki/Diagnostics-Guide.md)
- [Refactor Plan](./REFACTOR-PLAN.md) — the v3 → v4 findings list

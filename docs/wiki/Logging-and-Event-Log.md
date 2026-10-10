# Logging and Event Log

Every run produces a durable record: who ran it, against what, with which filters, what happened to
each file and each service, and the exit code. When something goes wrong on a print server at 3 a.m.,
the log is the only witness.

---

## Where the log goes

The **first writable** location wins:

| Priority | Location |
|---|---|
| 1 | `-LogPath <file>` (parent directory is created) |
| 2 | `%ProgramData%\PrintSpoolerCleanup\Logs\PrintSpoolerCleanup.log` |
| 3 | `%LocalAppData%\PrintSpoolerCleanup\Logs\…` |
| 4 | `<script directory>\Logs\…` |
| 5 | `%TEMP%\PrintSpoolerCleanup\Logs\…` |

The resolved path is printed in the execution summary and recorded in the report as `LogPath`.

> **Why not next to the script, like v3?** Because the script often lives somewhere read-only:
> `C:\Program Files`, a locked-down SMB share, a mounted ISO, a deployment cache. v3 wrote its log
> beside the script, every `Add-Content` failed, and the run silently degraded to console-only —
> exactly when you most needed a record.

Writability is proven with a **real open**, not an ACL guess. The probe writes no bytes, so it cannot
leave a blank first line that would break a strict line-oriented parser.

If every candidate fails you get one warning and the run continues without a file log:

```
WARNING: Logging disabled: cannot write to 'C:\...' (Access to the path is denied).
```

A logging failure never aborts a cleanup that is halfway through having the Spooler stopped.

### File naming

`Logging.SessionStampedFileName` controls the name:

| Value | File | Best for |
|---|---|---|
| `$false` *(default)* | `PrintSpoolerCleanup.log` + rotation | Scheduled tasks — one growing, rotating, prunable file |
| `$true` | `PrintSpoolerCleanup_20261010-031500.log` | Interactive use — one file per investigation |

`-Diagnostic`, `-Help`, `-ShowVersion`, `-CheckUpdate` and `-SelfTest` never create a session-stamped
file, because they produce no operational record worth keeping separately.

---

## Record format

Pipe-delimited, UTC, locale-independent:

```
2026-10-10T03:15:04.812Z|INFO    |Spooler|SRV01|Print Spooler stopped successfully
2026-10-10T03:15:06.440Z|WARN    |Cleanup|SRV01|Could not delete 'F00412.SHD' (locked=True, attempts=4): The process cannot access the file
2026-10-10T03:15:09.001Z|SUCCESS |Session|-|Run finished. Status=Success; targets=1; jobsRemoved=7; exitCode=0.
```

| Field | Notes |
|---|---|
| `ts` | UTC, ISO-8601 with milliseconds. Never local time, never a locale-dependent format |
| `level` | Left-padded to exactly 8 characters, so columns align in `grep` output |
| `component` | `Session`, `Spooler`, `Cleanup`, `Remote`, `Dispatch`, `Update`, `Config`, `Backup`, `Restore`, `Task`, `Watchdog`, `Security`, `EventLog`, `Report`, `Confirm`, `Platform`, `Payload`, `Diagnostics`, `SelfTest`, `Logging` |
| `target` | The computer being acted on, or `-` when not applicable |
| `message` | Newlines are flattened to spaces, so one record is always one line |

Sizes and durations inside a message use a **dot** as the decimal separator (`1.50 KB`, `1.5h`)
regardless of the host's regional settings. The process culture is pinned to `en-US` when the tool
starts, and the format helpers state the invariant culture explicitly as well, so a machine set to a
comma-separated locale still produces output that matches the documentation and stays greppable.

The last rule matters: a message containing a newline cannot forge extra records, which keeps the
file safe to parse line by line.

**Parsing in PowerShell:**

```powershell
Get-Content "$env:ProgramData\PrintSpoolerCleanup\Logs\PrintSpoolerCleanup.log" |
    ForEach-Object {
        $p = $_ -split '\|', 5
        [pscustomobject]@{ Time = [DateTime]::Parse($p[0]); Level = $p[1].Trim(); Component = $p[2]; Target = $p[3]; Message = $p[4] }
    } |
    Where-Object Level -in @('WARN', 'ERROR', 'CRITICAL')
```

**In bash:**

```bash
grep -E '\|(WARN|ERROR|CRITICAL)\s*\|' PrintSpoolerCleanup.log | cut -d'|' -f1,4,5
```

---

## JSON Lines mode

For log shippers (Fluentd, Filebeat, Vector, Azure Monitor agent):

```powershell
.\Clear-PrintSpoolerQueue.ps1 -JsonLog -Force
```

```json
{"ts":"2026-10-10T03:15:04.812Z","level":"INFO","component":"Spooler","target":"SRV01","msg":"Print Spooler stopped successfully"}
```

One object per line, no wrapping array, no trailing comma — directly consumable by anything that
speaks JSON Lines. A `data` property is added when the record carries structured detail.

---

## Levels

`TRACE` < `DEBUG` < `INFO` (= `SUCCESS`) < `WARN` < `ERROR` < `CRITICAL`

`-LogLevel` and `Logging.Level` set the **minimum** level recorded to file. Console rendering uses
the same threshold.

| Level | Use |
|---|---|
| `INFO` *(default)* | Normal operation. The right choice for scheduled tasks |
| `DEBUG` | Adds WinRM probe timings, configuration source, retention actions, update-check outcomes |
| `TRACE` | Adds skipped-update-check reasons and per-step detail. Verbose — use while investigating |
| `WARN` | Only problems. Useful when you ship logs centrally and want volume down |
| `ERROR` | Only failures |

`-Verbose` echoes every record to the verbose stream *in addition* to the file, which is the most
convenient way to watch a `-Repair` run live.

---

## Rotation and retention

**Rotation** is size-based:

```
PrintSpoolerCleanup.log      <- active
PrintSpoolerCleanup.log.1    <- most recent rotated
PrintSpoolerCleanup.log.2
...
PrintSpoolerCleanup.log.9    <- oldest; dropped on the next rotation
```

Controlled by `Logging.MaxSizeMB` (default 10) and `Logging.MaxFileCount` (default 9). Rotation runs
at the start of each run, so a long-lived scheduled task cannot grow a log without bound.
`-NoLogRotation` disables it for a single run.

**Retention** is age-based: `Logging.RetentionDays` (default 30). At the start of each run, files in
the log directory matching `*.log` or `*.log.<n>` and older than the window are deleted. The active
log file is never deleted.

Retention matches **any** `*.log` in the tool's own log directory rather than one hard-coded prefix,
because that directory is tool-owned and the tool itself produces several naming schemes (rotating,
session-stamped, custom `-LogPath`). An earlier version matched only the rotating name, so
session-stamped logs were never pruned and grew forever.

Set `RetentionDays = 0` to disable pruning — appropriate where log retention is governed by a policy
you manage elsewhere.

Rotation and retention are housekeeping: a failure in either is logged and the run continues.

---

## Thread safety

PowerShell 7 `-Parallel` workers log concurrently, so file writes are serialised with
`Monitor::Enter`/`Exit` on a dedicated lock object (never on `$this`, a string, or a type). The
in-memory ring buffer has its own lock for the same reason.

Without this, parallel fleet runs interleave partial lines and corrupt the log — which is worse than
not having one.

---

## The in-memory ring buffer

The last 1000 records are also kept in memory. This is what makes a diagnostics bundle possible when
the file log itself is unusable (read-only host, redirected `%TEMP%`, full disk):

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic -DiagnosticsBundlePath C:\Temp
# the bundle contains log-buffer.txt
```

---

## Windows Event Log

```powershell
.\Clear-PrintSpoolerQueue.ps1 -EventLog -Force
```

or `Logging.EventLog = $true` in the configuration file, or `PSQ_EVENT_LOG=true` machine-wide.

| Property | Value |
|---|---|
| Log | `Application` |
| Source | `PrintSpoolerCleanup` |
| Category | `Print Spooler Maintenance` |

| Event ID | Type | Written when |
|---:|---|---|
| 1000 | Information | An informational run completed |
| 1001 | Warning | `SuccessWithWarnings` or `PartialSuccess` |
| 1002 | Error | `Failed`, `ServiceFailure` |
| 1003 | SuccessAudit | A fully successful run |

The message body always includes the tool version and the log path, so an event is self-contained:

```
Print Spooler Queue Cleanup finished with status SuccessWithWarnings. Targets: 12 attempted,
10 succeeded, 2 failed. Jobs removed: 34. Files deleted: 68. Bytes reclaimed: 412.88 MB.
Exit code: 1.
Version : 4.0.0
Log     : C:\ProgramData\PrintSpoolerCleanup\Logs\PrintSpoolerCleanup.log
```

Creating the source requires elevation. On the first elevated run the source is created; if creation
fails (not elevated, or the Application log is locked down by policy) event logging is disabled **for
that run only**, a warning is logged, and the file log continues unaffected. Event Log integration is
an enhancement, never a dependency.

Query it:

```powershell
Get-WinEvent -FilterHashtable @{ LogName='Application'; ProviderName='PrintSpoolerCleanup' } -MaxEvents 50 |
    Format-Table TimeCreated, Id, LevelDisplayName, Message -Wrap
```

This is genuinely useful for a SIEM: a single Event Log query gives you the maintenance history of a
print server without shipping files anywhere.

> v3's README advertised Event Log integration that was never implemented. This is the real thing.

---

## What is logged

A normal run produces roughly:

```
Run started - Clear-PrintSpoolerQueue v4.0.0 - Elevated=True Invocation=f10bafad4e50 Log=... Mode=Cleanup PSVersion=7.6.6
Spool directory on 'SRV01': 'C:\Windows\System32\spool\PRINTERS'
Probing WinRM on 3 remote target(s)...
WinRM probe OK for 'SRV02' in 214 ms.
Processing the remote target 'SRV02' (1/3)...
Found 14 files (7 .SHD, 7 .SPL, 193273856 bytes) on 'SRV02'
Backup written: C:\ProgramData\...\spool-backup_SRV02_20261010-031500.zip
SHA-256 verified: 9f2c1a...
Print Spooler stopped successfully on 'SRV02'
Removed 14 of 14 spool files on 'SRV02'
Print Spooler is running on 'SRV02'
Run finished. Status=Success; targets=3; jobsRemoved=21; filesDeleted=42; reclaimed=412.88 MB; exitCode=0.
```

Plus, at `DEBUG` and above: configuration source, log retention actions, pending-update checks,
per-host probe timings.

**Never logged:** credentials, GitHub tokens, document contents, or the `Update.Token` configuration
value (redacted in the diagnostics bundle's `effective-config.json`).

---

## Related pages

- [Configuration](Configuration)
- [Reports and Machine-Readable Output](Reports-and-Machine-Readable-Output)
- [Security and Privacy](Security-and-Privacy)
- [Troubleshooting → Logging and output](Troubleshooting#logging-and-output)

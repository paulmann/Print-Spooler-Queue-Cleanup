# FAQ

The questions that come up in every rollout.

---

## Basics

### Do I need to install anything?

No. One `.ps1` file, run with Windows PowerShell 5.1 or PowerShell 7. No modules, no dependencies, no
internet access required to run. The `.bat` is optional and exists only so the tool can be
double-clicked.

### Is it safe to run on a production print server?

Yes, and it is designed for exactly that. `-Diagnostic` and `-WhatIf` are completely read-only — they
do not even stop the service. A real cleanup stops the Spooler for roughly 10–20 seconds per host.

If you want an archive first, `-Backup` writes one before anything is deleted. `-BackupRequired`
makes a failed backup abort the run.

### What exactly does it delete?

Only `*.SPL` and `*.SHD` files inside the resolved spool directory, and only after four independent
guards pass on that directory. Never anything else, never recursively, never on a UNC path.

The full list of refusals is in [Safety Model](Safety-Model).

### Will it delete my printers, ports or drivers?

No. Printers and ports are never touched. Drivers are only removed by
`-Repair -PurgeDriverStore -Force`, and then only packages that **no** printer references.

### Does it work on Windows Server?

Yes — Server 2012 R2 and later. The payload uses only cmdlets available in Windows PowerShell 5.1, so
the same code runs against an old Server target and a Windows 11 host with PowerShell 7.6.

On Server Core without the Print Services role there is no Spooler service, and the tool reports exit
code **10** (`UnsupportedPlatform`) rather than throwing a confusing "service not found".

### Does it work on Linux or macOS?

No, and it says so with exit code 10. The tool manages the Windows Print Spooler.

`-Help`, `-ShowVersion` and `-CheckUpdate` do run cross-platform, and the test suite runs on Linux in
CI — which is how a platform-dependent bug in the safety guards was found.

---

## Choosing options

### Should I use `-Force`?

For anything unattended, yes — without it the tool prompts, and a prompt nobody can answer means exit
code 4. Interactively, the prompt is worth keeping: it shows you whether you are about to delete
**ALL** spool files or **the SELECTED** ones.

### Should I use `-Backup`?

On a print server, yes, essentially always. On a workstation where you have just watched the queue
with `-Diagnostic`, it is optional.

The default location is `%ProgramData%\PrintSpoolerCleanup\Backups` on the target host — nothing
leaves the machine unless you point `-BackupPath` or `-ArchivePath` at a share.

### When should I use `-OlderThan`?

On any **shared** queue. A print server with 40 jobs usually has one wedged job and 39 that are fine;
deleting all 40 closes your ticket and opens thirty more.

```powershell
.\Clear-PrintSpoolerQueue.ps1 -OlderThan 2h -Force -Backup
```

It is also the right site-wide default in the config file, so automated runs can never destroy a job
that was submitted five minutes ago.

### `-Diagnostic` or `-WhatIf`?

Different questions.

| | `-Diagnostic` | `-WhatIf` |
|---|---|---|
| Question answered | *Why* is the queue stuck? | *What would this run do?* |
| Collects | Services, processes, printers, drivers, ports, event logs, jobs | The same inventory, plus the exact cleanup plan |
| Stops the service | No | No |
| Exit code | 11 | 0 |

Run `-Diagnostic` first when you do not know the cause. Run `-WhatIf` when you know and want to check
your filters before committing.

### `-Repair` or a normal run?

A normal run is the default and is enough most of the time. Use `-Repair` when:

- the Spooler will not stop (it cycles `PrintNotify`, `Fax` and `PrintWorkflowUserSvc_*` first)
- the same host keeps re-sticking
- you have seen repeated SCM 7031 crash events

`-Repair` is a **superset** — it does the whole cleanup plus the repair steps. You do not run both.

### Do I need `-ThrottleLimit`?

Only for fleets, and only deliberately. The default is **1** (sequential) because a Spooler restart
interrupts printing on that host for several seconds; running 64 at once during business hours is a
self-inflicted outage that looks like maintenance.

Above 1 needs PowerShell 7 for `ForEach-Object -Parallel`. On 5.1 a runspace pool is used
automatically, so the same command line works either way.

---

## Behaviour

### Why did it say "Nothing to do"?

Either the spool directory was already empty, or your filters matched nothing. Both map to exit code
**0**, because for a scheduled maintenance run "the queue was already clear" is a success, not a
warning.

If you expected files, check the selection line in the execution plan:

```
Selection       : jobs matching: printers like 'HP*'
```

### Why were some jobs not selected although I filtered by printer?

When the Spooler is hung it cannot answer WMI, so the inventory comes from `.SHD` headers where the
printer name is best-effort. Jobs with no readable metadata are **deliberately not selected** by
name-based filters, and the run reports this as a warning.

That is a choice, not a limitation: silently guessing which printer a job belongs to and then
deleting it would be worse than doing nothing. Use `-JobId` (always derived from the file name) or
`-OlderThan` (uses file timestamps) instead.

### The run succeeded but files reappeared

`ResurrectedFiles > 0` in the report. A driver or port monitor is re-submitting them, or a client
keeps resending. Find the submitter before clearing again:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic | Select-Object -ExpandProperty Results |
    Select-Object -ExpandProperty Jobs | Format-Table JobId, PrinterName, UserName, DocumentName, Submitted
```

### Why is exit code 11 not an error?

`-Diagnostic` completed. It is informational — you asked for a report, you got one. Treat 0 and 11 as
healthy in monitoring.

### It reported "logging disabled"

Every candidate log location was unwritable. The run continued; only the file record was lost, and the
console output and report were unaffected. Pass `-LogPath` to force a location you know is writable.

### Can I run two cleanups at once?

On the same host, no — the second will find the Spooler already stopped and files already gone, and
`MultipleInstancesPolicy = IgnoreNew` in the scheduled task prevents overlap by design. On different
hosts, that is exactly what `-ThrottleLimit` is for.

### Does it restart the Fax service?

Only if Fax was **running** when the run started. A stopped or disabled Fax service is left exactly as
it was found. A cleanup must never silently change a host's configuration.

### What if the Spooler cannot be stopped?

**Nothing is deleted**, the run returns `ServiceFailure` (exit **12**), dependents are restarted, and
a last-resort attempt is made to bring the Spooler back up. The service is left as close to its
original state as possible.

That ordering is deliberate: deleting files while `spoolsv.exe` still holds the job table in memory
produces an inconsistent queue, which is worse than a queue that is still stuck.

---

## Restoring

### I deleted something I needed. Can I get it back?

If you ran with `-Backup`, the document payload is in the archive:

```powershell
Get-SpoolerBackup | Format-Table CreatedUtc, ComputerName, FileCount, SizeBytes
Restore-SpoolerBackup -Path <archive.zip> -DestinationDirectory C:\Temp\restored
```

If you ran without `-Backup`, the files are gone. There is no undelete.

### Can I put the files back and have the job print?

**No.** `spoolsv.exe` keeps its job table in memory and only reads `.SHD`/`.SPL` for jobs it created
in the current boot. Files copied back into `PRINTERS` are ignored or reaped.

A restore is a **forensic** recovery: you get the rendered document back so it can be re-submitted
from the application. For RAW spool data you can often send the `.SPL` straight to the device; for EMF
the document must be re-rendered. The manifest's `DocumentName` and `Owner` tell you who to ask.

The tool says this in four separate places on purpose. A runbook that promises a rollback which cannot
happen gets someone paged at 3 a.m.

---

## Automation

### Will a scheduled task block on a prompt?

No. Tasks created by `Install-SpoolerCleanupTask` always pass `-Force -NonInteractive -Quiet
-NoProgress`. And independently, the tool detects a non-interactive host
(`[Console]::IsInputRedirected`) and refuses to prompt — it exits **4** rather than hanging.

That second layer exists because the v3 README's `schtasks` one-liner ran the interactive `.bat` as
SYSTEM, and "Press any key" blocked the task until its time limit killed it.

### Can I use the `.bat` in a scheduled task?

You can, with `/Force /Quiet /NoPause`. But you should not: the `.ps1` gives you the same exit codes
with less indirection, and the installer already builds the right command line.

### How do I feed results into my RMM?

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Force -Backup -NonInteractive -Quiet -NoProgress -MachineReadableJson
```

Exactly one JSON document on stdout, nothing else — no banner, no progress, no colour escapes. The
process exit code carries the outcome. Schema in
[Reports and Machine-Readable Output](Reports-and-Machine-Readable-Output).

### Can I poll health without cleaning?

Yes — `Test-PrintSpoolerHealth`. Non-destructive, returns a 0–100 score with plain-English findings,
safe to poll frequently:

```powershell
Get-Content hosts.txt | Test-PrintSpoolerHealth -ThrottleLimit 16 | Where-Object Health -ne 'Healthy'
```

---

## Updates

### Does it phone home?

Only if you ask it to. `-CheckUpdate`, `-Update` or `-AutoUpdate` contact GitHub. A normal cleanup
makes **no** network calls other than WinRM to the targets you named.

Tasks created by `Install-SpoolerCleanupTask` include `-AutoUpdate`, which checks at most once per
cooldown window (24 h default), never blocks, and never fails the run. Disable it with
`-NoAutoUpdate` or `PSQ_NO_UPDATE_CHECK=true`.

### What if GitHub is unreachable?

Exit code **6** for `-CheckUpdate`, and a `DEBUG`-level log entry for `-AutoUpdate`. The cleanup is
unaffected. Air-gapped sites should set `PSQ_NO_UPDATE_CHECK=true` so it stops appearing in logs.

### Can it install a tampered update?

Not by default. Downloads are verified against the SHA-256 published in the release's
`checksums.txt`; on mismatch the download is discarded and the installed copy is untouched. When a
release publishes no checksums the update is **refused** unless you pass
`-SkipChecksumVerification` explicitly.

Downloaded scripts are also parsed before being accepted, so a truncated download or a proxy returning
an HTML error page is rejected.

### Can an update break the running tool?

No. A process never overwrites the file it is executing from. Updates are staged in `.update\` with a
per-file SHA-256 manifest and swapped in at the **start of the next run**, after re-verification. The
previous version is kept in `backup-<version>-<stamp>\`.

---

## Troubleshooting shortcuts

| Symptom | First command |
|---|---|
| Queue stuck, cause unknown | `-Diagnostic` |
| `net stop spooler` hangs | `-Repair -Force -Backup` |
| Files come straight back | `-Diagnostic`, then look at `ResurrectedFiles` and the driver/port columns |
| Remote targets all time out | `-PreFlightProbe` (on by default) — read the per-host warning |
| `-PrinterName` matched nothing | `-Diagnostic` and check the `Source` column: `SHD` means best-effort metadata |
| Task runs but does nothing | `Get-SpoolerCleanupTask` → `LastTaskResultMeaning` |
| Update check fails | Exit 6; check proxy, or set `PSQ_GITHUB_TOKEN` / `PSQ_NO_UPDATE_CHECK` |
| Garbled characters in cmd | `/Ascii`, or `chcp 65001` |
| Not sure the copy is sound | `-SelfTest` |

Full matrix: [Troubleshooting](Troubleshooting).

---

## Licensing and support

### What is the license?

**GPL-3.0-or-later.** See
[LICENSE](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/blob/main/LICENSE).

Earlier revisions of the README displayed a "Public Domain" badge while the `LICENSE` file contained
the GPL v3 text. The badge was wrong; the license file is authoritative and the README now matches it.

### Can I use it commercially?

Yes. GPL-3.0 permits commercial use. If you distribute a modified version, the GPL's copyleft terms
apply to your derivative.

Internal use — running it on your own machines, from your own share, in your own scheduled tasks —
involves no distribution and triggers no copyleft obligation.

### Can I modify it?

Yes. `src/` is the source of truth; `tools/Build.ps1` assembles the distributables. See
[For Developers](For-Developers).

### How do I report a bug?

Produce a bundle and
[open an issue](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/issues):

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic -DiagnosticsBundlePath C:\Temp
```

Include the bundle, the exact command line, the exit code, and what you expected. The bundle contains
the environment, diagnostics, effective configuration **with secrets redacted**, the recent log buffer
and file integrity hashes — never credentials, tokens or document contents.

For a **security** issue, email Mikhail@Deynekin.com instead of opening a public issue.

### Who maintains it?

Mikhail Deynekin — <Mikhail@Deynekin.com> — [Deynekin.com](https://Deynekin.com)

---

## Related pages

- [Getting Started](Getting-Started)
- [CLI Reference](CLI-Reference)
- [Troubleshooting](Troubleshooting)
- [Safety Model](Safety-Model)
- [Why Print Queues Get Stuck](Why-Print-Queues-Get-Stuck)

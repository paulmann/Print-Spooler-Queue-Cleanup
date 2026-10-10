# 🖨️ Print Spooler Queue Cleanup

[![License: GPL v3](https://img.shields.io/badge/License-GPL--3.0--or--later-blue.svg)](./LICENSE)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207.x-blue.svg)](https://docs.microsoft.com/powershell/)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%20%7C%2011%20%7C%20Server-0078d4.svg)](https://www.microsoft.com/windows)
[![Version](https://img.shields.io/badge/version-4.0.1-green.svg)](./CHANGELOG.md)
[![Tests](https://img.shields.io/badge/tests-715%20passing-brightgreen.svg)](./tests/Run-Tests.ps1)

**Clears stuck Windows print queues safely, selectively and verifiably — locally or across a
fleet — with backup, diagnostics, repair, structured logging, machine-readable reports and
built-in self-updating.**

One PowerShell file. No modules to install, no dependencies, no internet access required to run.

---

## The problem

The Print Spooler keeps queued jobs as `*.SHD` (metadata) and `*.SPL` (data) files in
`%SystemRoot%\System32\spool\PRINTERS`. When a job cannot be rendered, a driver faults, or a
port monitor blocks, the queue **wedges**: the service reports `Running`, jobs will not cancel,
and everything behind them stalls.

The canonical fix is three commands:

```cmd
net stop spooler
del /F /Q "%SystemRoot%\System32\spool\PRINTERS\*.*"
net start spooler
```

It is correct — and in production it fails constantly, because:

| Reality | Consequence of the naive fix |
|---|---|
| `Fax`, `PrintNotify` and `PrintWorkflowUserSvc_*` hold handles | `net stop` hangs in `STOP_PENDING` |
| Files are locked by `spoolsv.exe` | The first locked file aborts the run |
| `del *.*` matches everything | Non-spool files in the folder are destroyed |
| The queue holds 40 healthy jobs | 40 users' work is deleted to fix one bad job |
| No archive is taken | There is no undo |
| Nothing is verified afterwards | "Service is Running" ≠ "printing works" |

This tool exists to close each of those gaps.

---

## Quick start

```powershell
# 1. See what would happen. Changes nothing, does not even stop the service.
.\Clear-PrintSpoolerQueue.ps1 -WhatIf

# 2. Look at why the queue is stuck (read-only deep inspection).
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic

# 3. Fix it, archiving everything first.
.\Clear-PrintSpoolerQueue.ps1 -Force -Backup

# 4. Read the manual without leaving the console.
.\Clear-PrintSpoolerQueue.ps1 -Help
```

Double-click users: run **`Clear-PrintSpoolerQueue.bat`**. It finds PowerShell 7 (or falls back
to 5.1), can self-elevate through UAC with `/Elevate`, and never pauses when launched by Task
Scheduler.

---

## What is new in v4.0.0

**v4.0.1** is a maintenance release on top of it: it fixes the CI workflow that never ran, the
self-update version check, and locale-dependent number formatting. See [CHANGELOG.md](./CHANGELOG.md)
for the details.

v3.1.0 stopped the Spooler, deleted `*.SHD`/`*.SPL`, and restarted it. That is still the core,
but it was the only thing it could do. v4 is a rebuild around what actually happens on print
servers:

**Precision** — act on *some* jobs instead of all of them:
`-PrinterName`, `-ExcludePrinterName`, `-JobId`, `-OwnerName`, `-DocumentName`,
`-OlderThan`, `-NewerThan`, `-LargerThanBytes`, `-SmallerThanBytes`.
Jobs are inventoried from `Win32_PrintJob`, with a `.SHD`-header fallback for when the Spooler is
too hung to answer WMI — which is exactly when you are running this tool.

**Safety** — `-Backup` archives every file it is about to delete (ZIP + JSON manifest + SHA-256),
with `Restore-SpoolerBackup` to recover them. Four independent guards fence the delete step, and
`-BackupRequired` turns a failed backup into a hard abort.

**Root cause** — `-Diagnostic` reports services, `spoolsv.exe` health, printers, drivers, ports,
spool-directory contents, the job table and the relevant event logs. `-Repair` cycles dependent
services, recovers `PrintNotify`, purges the Forms cache, and can remove unreferenced driver
packages or reset a corrupt `DefaultSpoolDirectory`.

**Scale** — `-ThrottleLimit N` runs a fleet in parallel (PowerShell 7 `ForEach-Object -Parallel`,
or a runspace pool on 5.1), `-PreFlightProbe` finds offline hosts in about a second instead of 40,
and one failing computer can never abort the rest.

**Observability** — rotating size- and age-capped logs in a pipe-delimited UTC format (or JSON
Lines), optional Windows Event Log integration, HTML/CSV/JSON/Text reports, and
`-MachineReadableJson` which puts exactly one JSON document on stdout for RMM ingestion.

**Lifecycle** — `-CheckUpdate` / `-Update` / `-AutoUpdate` track GitHub Releases with SHA-256
verification and staged self-update; `Install-SpoolerCleanupTask` registers a maintenance task
that cannot block on a prompt.

**Answers** — `-Help` renders a complete reference (modes, every parameter, 16 worked examples,
exit codes, config schema, environment variables, log and report formats, WinRM prerequisites,
scheduling, update mechanics, the safety model and a troubleshooting matrix).

---

## Three ways to run it

| Form | Use when | Example |
|---|---|---|
| **`Clear-PrintSpoolerQueue.ps1`** | You have a shell. This is the real tool. | `.\Clear-PrintSpoolerQueue.ps1 -Force -Backup` |
| **`Clear-PrintSpoolerQueue.bat`** | Double-click from Explorer, or handing a file to a technician. | `Clear-PrintSpoolerQueue.bat /Force /Backup` |
| **`PrintSpoolerCleanup` module** | You want discoverable cmdlets in automation. | `Import-Module .\PrintSpoolerCleanup` |

The launcher is deliberately thin: it resolves an engine, elevates, applies a staged update,
translates arguments and propagates the exit code. **There is exactly one implementation of the
destructive step**, so batch and PowerShell cannot disagree about what gets deleted.

---

## Common operations

```powershell
# Only jobs older than two hours — the busy-queue-safe default
.\Clear-PrintSpoolerQueue.ps1 -OlderThan 2h -Force -Backup

# One printer on one server
.\Clear-PrintSpoolerQueue.ps1 -ComputerName SRV-PRINT01 -PrinterName "KYOCERA*" -Force

# Specific job ids
.\Clear-PrintSpoolerQueue.ps1 -JobId 4217,4218 -Force -Backup

# A whole site, in parallel, with a CSV for the report
.\Clear-PrintSpoolerQueue.ps1 -ComputerName (Get-Content .\hosts.txt) -ThrottleLimit 24 `
    -Force -Backup -ExportFormat Csv -ExportPath \\filesrv\logs\spooler

# A machine that keeps re-sticking
.\Clear-PrintSpoolerQueue.ps1 -Repair -PurgeFormsCache -Force -Backup -Verbose

# Unattended, machine-readable
.\Clear-PrintSpoolerQueue.ps1 -Force -NonInteractive -Quiet -MachineReadableJson -OlderThan 6h -Backup

# Nightly maintenance
Install-SpoolerCleanupTask -ScriptPath C:\Tools\Clear-PrintSpoolerQueue.ps1 -Schedule Daily -StartTime 03:00

# Self-heal after a spooler crash (SCM event 7031)
Install-SpoolerCleanupTask -ScriptPath C:\Tools\Clear-PrintSpoolerQueue.ps1 -Schedule OnEvent `
    -EventLogName System -EventSource "Service Control Manager" -EventId 7031

# Health probe for monitoring
Test-PrintSpoolerHealth -ComputerName SRV-PRINT01 | ConvertTo-Json

# Is there a newer version?
.\Clear-PrintSpoolerQueue.ps1 -CheckUpdate      # exit code 7 = yes
```

---

## Exit codes

Identical for the `.ps1` and the `.bat`, so a monitoring rule written for one works for the other.

| Code | Meaning | | Code | Meaning |
|---:|---|---|---:|---|
| 0 | Success | | 8 | Update applied |
| 1 | Partial success | | 9 | Update failed |
| 2 | Failed | | 10 | Unsupported platform |
| 3 | Bad arguments | | 11 | Diagnostics only *(informational)* |
| 4 | Cancelled by user | | 12 | Spooler service failure |
| 5 | Not administrator | | 13 | Restore failed |
| 6 | Update check failed | | 14 | Signature invalid |
| 7 | Update available | | 15 | Watchdog timeout |

Alert on **2, 5, 12, 15**. Treat **1** as "needs attention", **7** as "plan an update",
and **0 / 11** as healthy.

---

## Requirements

- Windows 10 / 11 / Server 2012 R2 or later, with the Print Spooler feature installed
- Windows PowerShell **5.1** or PowerShell **7.x** (7.x required for `-ThrottleLimit` > 1)
- **Administrator** rights for the local machine
- **WinRM** plus authorization for remote targets (`Enable-PSRemoting -Force`)
- No external modules, no internet access needed to run

---

## Installation

```powershell
# Option A: download the two files
Invoke-WebRequest https://github.com/paulmann/Print-Spooler-Queue-Cleanup/releases/latest/download/Clear-PrintSpoolerQueue.ps1 -OutFile Clear-PrintSpoolerQueue.ps1
Invoke-WebRequest https://github.com/paulmann/Print-Spooler-Queue-Cleanup/releases/latest/download/Clear-PrintSpoolerQueue.bat -OutFile Clear-PrintSpoolerQueue.bat
Unblock-File .\Clear-PrintSpoolerQueue.ps1

# Option B: clone
git clone https://github.com/paulmann/Print-Spooler-Queue-Cleanup.git

# Option C: enterprise share + scheduled task
Copy-Item .\Clear-PrintSpoolerQueue.ps1 \\server\share\Scripts\PrintCleanup\
Install-SpoolerCleanupTask -ScriptPath \\server\share\Scripts\PrintCleanup\Clear-PrintSpoolerQueue.ps1
```

Verify what you downloaded:

```powershell
Get-FileHash .\Clear-PrintSpoolerQueue.ps1 -Algorithm SHA256
# compare with checksums.txt on the release page
```

---

## Documentation

The built-in reference is the most complete and is always in sync with the code:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Help              # everything
.\Clear-PrintSpoolerQueue.ps1 -HelpSection Safety    # one section
.\Clear-PrintSpoolerQueue.ps1 -ListHelpSections
```

The [**wiki**](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki) covers the same
ground in depth, plus the background material:

| | |
|---|---|
| [Getting Started](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/Getting-Started) | [CLI Reference](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/CLI-Reference) |
| [Installation & Deployment](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/Installation-and-Deployment) | [Cmdlet Reference](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/Cmdlet-Reference) |
| [Print Spooler Internals](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/Print-Spooler-Internals) | [Configuration](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/Configuration) |
| [Why Queues Get Stuck](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/Why-Print-Queues-Get-Stuck) | [Logging & Event Log](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/Logging-and-Event-Log) |
| [Safety Model](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/Safety-Model) | [Reports & Machine-Readable Output](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/Reports-and-Machine-Readable-Output) |
| [Exit Codes](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/Exit-Codes) | [Remote & Fleet Management](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/Remote-and-Fleet-Management) |
| [Backup & Restore](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/Backup-and-Restore) | [Scheduling & Automation](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/Scheduling-and-Automation) |
| [Diagnostics Guide](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/Diagnostics-Guide) | [Updates & Versioning](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/Updates-and-Versioning) |
| [Repair & Deep Cleaning](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/Repair-and-Deep-Cleaning) | [Troubleshooting](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/Troubleshooting) |
| [Enterprise Deployment](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/Enterprise-Deployment) | [Batch Launcher Reference](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/Batch-Launcher-Reference) |
| [For Developers](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/For-Developers) | [Security & Privacy](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/Security-and-Privacy) |
| [FAQ](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki/FAQ) | |

---

## Repository layout

```
Clear-PrintSpoolerQueue.ps1        standalone script  (generated - do not edit)
Clear-PrintSpoolerQueue.bat        launcher           (generated from bat/)
PrintSpoolerCleanup.config.psd1    documented configuration template
src/                               source of truth (Shared / Private / Public)
bat/                               launcher template
tools/Build.ps1                    assembles dist/ from src/ (runs the gate, packages dist/)
tools/Invoke-Validation.ps1        parse gate
tools/Invoke-CollisionScan.ps1     parameter/local collision detector (self-tested)
tools/Invoke-AutomaticVariableScan.ps1  automatic-variable write detector (self-tested)
tools/Invoke-PS51CompatibilityScan.ps1  PS7-only syntax detector for the 5.1 floor
tools/ps51-scanner-selftest.ps1      16 fixtures proving that detector fires AND stays quiet
tools/Invoke-WorkflowLint.ps1      parse-validates the PowerShell embedded in CI YAML
tools/Invoke-DocLint.ps1           structural validation of every Markdown document
tools/Invoke-TemplateLint.ps1      validates .github YAML (escapes, issue-form schema)
tools/template-lint-selftest.ps1   12 fixtures proving that linter fires AND stays quiet
tools/module-test.ps1              validates the .psd1/.psm1 form  (gate stage 6)
tools/smoke-test.ps1               end-to-end checks against the built artifact  (stage 7)
tools/Verify-Release.ps1           audits the shippable artifacts  (stage 8)
tools/Verify-Brief.ps1             proves every brief requirement still works  (stage 9)
tools/Publish-Wiki.ps1             publishes docs/wiki/ to the GitHub wiki
tests/                             harness + unit suites + nine-stage quality gate (715 checks)
docs/                              contracts and engineering notes
.github/workflows/                 CI
```

`src/` is the only place to make changes. `tools/Build.ps1` assembles both distributables from
it, so the script and the module cannot drift apart — which is exactly what happened to v3.

```powershell
./tools/Build.ps1        # build + validate + package + checksums
./tests/Run-Tests.ps1    # full quality gate
```

---

## License

**GPL-3.0-or-later** — see [LICENSE](./LICENSE).

> Earlier revisions of this README displayed a "Public Domain" badge while the `LICENSE` file
> contained the GPL v3 text. The badge was wrong; the license file is authoritative and this
> README now matches it.

---

## Author

**Mikhail Deynekin** — <Mikhail@Deynekin.com> — [Deynekin.com](https://Deynekin.com)

Issues, diagnostics bundles and pull requests: [github.com/paulmann/Print-Spooler-Queue-Cleanup](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/issues)

Built for senior system administrators who would rather close the ticket permanently than
restart the service every morning.

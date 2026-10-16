# Print Spooler Queue Cleanup — Wiki

**Clears stuck Windows print queues safely, selectively and verifiably — now as a native
Win32 app with a GUI, tray guardian, built-in installer and self-update.**

> **v5.0.0 (native)** · Windows 10 / 11 x64 · one static `SpoolerCleanup.exe`, no .NET, no
> redistributables · PolyForm Noncommercial 1.0.0
>
> Start here: [GUI Application](GUI-Application) · [Tray](Tray-Icon-and-Notifications) ·
> [Themes & Languages](Themes-and-Languages) · [Installer & Updates](Installer-Uninstall-and-Updates) ·
> [CLI & Automation](CLI-and-Automation)

The pages below the *PowerShell legacy (v4)* section document the v4 script kept under
`legacy/powershell/` — still the right tool for fleet/remote scenarios.

---

## Start here

**New to the tool?** Read [Getting Started](Getting-Started) — five minutes from download to
your first safe cleanup.

**Something is on fire right now?**

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic        # read-only: why is it stuck?
.\Clear-PrintSpoolerQueue.ps1 -Force -Backup     # fix it, archive everything first
```

**In a hurry and want the whole manual in your console?**

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Help              # 19 sections, always in sync with the code
.\Clear-PrintSpoolerQueue.ps1 -HelpSection Troubleshooting
```

---

## The problem in one paragraph

The Print Spooler keeps queued jobs as `*.SHD` (metadata) and `*.SPL` (data) files in
`%SystemRoot%\System32\spool\PRINTERS`. When a job cannot be rendered, a driver faults, or a port
monitor blocks, the queue wedges: the service reports `Running`, jobs will not cancel, and
everything queued behind them stalls. The textbook fix — stop the service, delete the folder
contents, start the service — is correct and still fails constantly in production, because
dependent services hold the stop open, files are locked, healthy jobs get destroyed, there is no
undo, and nobody verifies that printing works afterwards. This tool closes each of those gaps.
The full background is in [Why Print Queues Get Stuck](Why-Print-Queues-Get-Stuck) and
[Print Spooler Internals](Print-Spooler-Internals).

---

## Wiki contents

### Getting going
| Page | What it covers |
|---|---|
| [Getting Started](Getting-Started) | Download, first safe run, the four commands you will actually use |
| [Installation and Deployment](Installation-and-Deployment) | Standalone, module, network share, GPO/Intune/ConfigMgr packaging |
| [Batch Launcher Reference](Batch-Launcher-Reference) | Every `.bat` switch, engine selection, elevation, scheduling caveats |

### Reference
| Page | What it covers |
|---|---|
| [CLI Reference](CLI-Reference) | All 70+ parameters, grouped, with defaults and gotchas |
| [Cmdlet Reference](Cmdlet-Reference) | The 14 exported cmdlets of the `PrintSpoolerCleanup` module |
| [Configuration](Configuration) | The `.psd1` schema, `PSQ_*` variables, precedence, validation rules |
| [Exit Codes](Exit-Codes) | The 0–15 contract and how to build monitoring rules on it |
| [Logging and Event Log](Logging-and-Event-Log) | File format, JSON Lines, rotation, retention, Event Log source |
| [Reports and Machine-Readable Output](Reports-and-Machine-Readable-Output) | JSON/CSV/HTML/Text schemas and RMM ingestion |

### Operations
| Page | What it covers |
|---|---|
| [Diagnostics Guide](Diagnostics-Guide) | Reading `-Diagnostic` output; the six patterns that explain 90% of tickets |
| [Repair and Deep Cleaning](Repair-and-Deep-Cleaning) | `-Repair`, Forms cache, driver store, spool-directory reset |
| [Backup and Restore](Backup-and-Restore) | What a backup contains, and the honest limits of restoring spool files |
| [Remote and Fleet Management](Remote-and-Fleet-Management) | WinRM prerequisites, parallelism, pre-flight probing, credentials |
| [Scheduling and Automation](Scheduling-and-Automation) | Task Scheduler, event-driven self-healing, CI and RMM integration |
| [Updates and Versioning](Updates-and-Versioning) | Discovery strategies, checksum verification, staged self-update |
| [Enterprise Deployment](Enterprise-Deployment) | Site-wide rollout, policy, monitoring, reporting at scale |
| [Troubleshooting](Troubleshooting) | Symptom → cause → fix, with the exact command for each |
| [FAQ](FAQ) | The questions that come up in every rollout |

### Background and governance
| Page | What it covers |
|---|---|
| [Print Spooler Internals](Print-Spooler-Internals) | SHD/SPL files, the job table, dependent services, why the stop hangs |
| [Why Print Queues Get Stuck](Why-Print-Queues-Get-Stuck) | The seven root causes and how each one presents |
| [Safety Model](Safety-Model) | The four delete guards and everything the tool refuses to do |
| [Security and Privacy](Security-and-Privacy) | Signature verification, config sandboxing, secret handling, telemetry |
| [For Developers](For-Developers) | Repository layout, build pipeline, test suite, contributing |

---

## Thirty-second tour

```powershell
# What is wrong? (read-only, safe on production at any time)
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic -ExportFormat Html -ExportPath C:\Temp

# Fix only the wedged jobs, keep the healthy ones, archive first
.\Clear-PrintSpoolerQueue.ps1 -OlderThan 2h -Force -Backup

# A site of 400 workstations, 24 at a time
.\Clear-PrintSpoolerQueue.ps1 -ComputerName (Get-Content .\hosts.txt) -ThrottleLimit 24 -Force -Backup

# A machine that keeps re-sticking
.\Clear-PrintSpoolerQueue.ps1 -Repair -PurgeFormsCache -Force -Backup

# Nightly maintenance that can never block on a prompt
Install-SpoolerCleanupTask -ScriptPath C:\Tools\Clear-PrintSpoolerQueue.ps1 -Schedule Daily -StartTime 03:00

# Self-heal within a minute of a spooler crash
Install-SpoolerCleanupTask -ScriptPath C:\Tools\Clear-PrintSpoolerQueue.ps1 -Schedule OnEvent `
    -EventLogName System -EventSource "Service Control Manager" -EventId 7031
```

---

## Design principles

These are the rules the project is held to, and the reason it looks the way it does.

1. **One destructive implementation.** The delete step exists exactly once, as a payload that runs
   on the target machine. Local and remote execution use the same code, so they cannot drift. The
   batch launcher delegates rather than reimplementing — v3 had two copies and they disagreed.
2. **Guards over good intentions.** Four independent checks fence the delete step. If any fails,
   nothing is deleted. The tool would rather do nothing than delete the wrong directory.
3. **Never leave a host without printing.** Every failure path makes a last-resort attempt to
   restart the Spooler. A failed cleanup must not become an outage.
4. **Prove it, do not assume it.** The service state is re-read after every transition; the queue
   is re-scanned after restart; `-SelfTestPrint` submits a real job. "Running" is not the same as
   "working".
5. **Be honest about limits.** Restoring spool files does not re-queue jobs, and the docs say so
   instead of implying a rollback that cannot happen.
6. **Fail loudly on configuration.** An unknown config key or an out-of-range value aborts the run
   with exit 3. A silently ignored setting is how a site loses its backups.
7. **Machine-readable by default.** Every run can produce a stable JSON contract and a documented
   exit code, so the tool is a citizen of your monitoring stack rather than a script someone runs.

---

## License and support

GPL-3.0-or-later — see [LICENSE](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/blob/main/LICENSE).

Bugs and feature requests: [open an issue](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/issues).
Please attach a diagnostics bundle — it makes a fix possible in one round trip instead of five:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic -DiagnosticsBundlePath C:\Temp
```

The bundle contains the environment report, diagnostics, effective configuration with secrets
redacted, the recent log buffer and file integrity hashes. It never contains credentials or
document contents.

**Author:** Mikhail Deynekin — <Mikhail@Deynekin.com> — [Deynekin.com](https://Deynekin.com)

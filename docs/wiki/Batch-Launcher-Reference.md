# Batch Launcher Reference

`Clear-PrintSpoolerQueue.bat` exists for two situations where a batch file is genuinely the right
answer: double-clicking from Explorer, and handing a technician something that runs in a bare cmd
prompt.

Everything substantive happens in PowerShell. The launcher only resolves an engine, optionally
elevates, applies a staged update, translates arguments and propagates the exit code.

---

## Why it is thin

v3.1.0 reimplemented the entire cleanup in batch. That meant **two implementations of one destructive
operation** — two chances to be wrong, and a batch copy that could never offer remote, selective or
backed-up cleanup because none of those fit in `cmd.exe`.

v4 delegates. There is exactly one place where a file can be deleted, so "what can this tool delete?"
has one auditable answer. The launcher contains no `del`, no `net stop` and no service manipulation,
and a build-time test asserts that.

---

## Basic use

```cmd
Clear-PrintSpoolerQueue.bat /?                      rem help
Clear-PrintSpoolerQueue.bat /Diagnostic             rem read-only
Clear-PrintSpoolerQueue.bat /Force /Backup          rem fix it
Clear-PrintSpoolerQueue.bat /Elevate                rem request UAC first
Clear-PrintSpoolerQueue.bat /Version
Clear-PrintSpoolerQueue.bat /Help Safety            rem a section of the full reference
```

---

## Engine selection

| Order | Engine | Notes |
|---|---|---|
| 1 | `pwsh.exe` (PowerShell 7) | First match on `PATH`. Required for `-ThrottleLimit > 1` |
| 2 | `%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe` | Always present on Windows |

Override:

| Switch / variable | Effect |
|---|---|
| `/PS7` | Require `pwsh.exe`; fail with exit 10 if absent |
| `/PS5` | Require Windows PowerShell 5.1 |
| `set PSQ_ENGINE=pwsh` | Same as `/PS7` |
| `set PSQ_ENGINE=powershell` | Same as `/PS5` |

The launcher reports which engine it chose unless `/Quiet` is given.

---

## Complete switch list

### Modes

| Switch | PowerShell equivalent |
|---|---|
| `/Diagnostic` | `-Diagnostic` |
| `/WhatIf`, `/Dry-Run` | `-WhatIf` |
| `/Repair` | `-Repair` |
| `/Restore <zip>` | `-Restore <zip>` |
| `/RestoreDir <dir>` | `-RestoreDestination <dir>` |
| `/CheckUpdate` | `-CheckUpdate` |
| `/Update` | `-Update -Force` |
| `/AutoUpdate` | `-AutoUpdate` |
| `/SelfTest` | `-SelfTest` |
| `/SelfTestPrint` | `-SelfTestPrint` |
| `/Help [<section>]` | `-HelpSection <section>` |
| `/?`, `/h`, `--help` | local help screen |
| `/Version`, `-v` | version |

### Selection

| Switch | PowerShell equivalent |
|---|---|
| `/Printer <name>` | `-PrinterName <name>` |
| `/ExcludePrinter <name>` | `-ExcludePrinterName <name>` |
| `/OlderThan <age>` | `-OlderThan <age>` |
| `/NewerThan <age>` | `-NewerThan <age>` |
| `/JobId <ids>` | `-JobId <ids>` |
| `/Owner <name>` | `-OwnerName <name>` |
| `/Computers <list>` | `-ComputerName <list>` |
| `/Hosts <file>` | `-ComputerListFile <file>` |
| `/Throttle <n>` | `-ThrottleLimit <n>` |

### Safety

| Switch | PowerShell equivalent |
|---|---|
| `/Force`, `/Y` | `-Force` |
| `/Backup` | `-Backup` |
| `/BackupPath <dir>` | `-BackupPath <dir>` |
| `/ArchivePath <dir>` | `-ArchivePath <dir>` |
| `/Elevate` | `-SelfElevate` |
| `/NonInteractive` | `-NonInteractive` |
| `/VerifySignature` | `-VerifySignature` |
| `/Watchdog` | `-Watchdog` |
| `/Timeout <sec>` | `-ServiceTimeoutSeconds <sec>` |
| `/Settle <sec>` | `-SettleSeconds <sec>` |

### Repair options

| Switch | PowerShell equivalent |
|---|---|
| `/PurgeForms` | `-PurgeFormsCache` |
| `/PurgeDrivers` | `-PurgeDriverStore` |
| `/ResetSpool` | `-ResetSpoolDirectory` |

### Output

| Switch | PowerShell equivalent |
|---|---|
| `/Quiet` | `-Quiet -NoProgress` |
| `/NoPause` | *(launcher-only: do not wait for a keypress)* |
| `/NoColor` | `-NoColor` |
| `/Ascii` | `-Ascii` |
| `/Json` | `-MachineReadableJson -Quiet -NoProgress` |
| `/JsonLog` | `-JsonLog` |
| `/Log <file>` | `-LogPath <file>` |
| `/LogLevel <lvl>` | `-LogLevel <lvl>` |
| `/RetentionDays <n>` | `-LogRetentionDays <n>` |
| `/ExportFormat <f>` | `-ExportFormat <f>` |
| `/ExportPath <dir>` | `-ExportPath <dir>` |
| `/Bundle <dir>` | `-DiagnosticsBundlePath <dir>` |
| `/EventLog` | `-EventLog` |

### Configuration and updates

| Switch | PowerShell equivalent |
|---|---|
| `/Config <file>` | `-ConfigPath <file>` |
| `/SaveConfig <file>` | `-SaveConfig -ConfigPath <file>` |
| `/Channel <name>` | `-UpdateChannel <name>` |

Unknown arguments exit **3** with a pointer to `/?`.

---

## Exit codes

Identical to the `.ps1`, so a monitoring rule written for one works for the other:

```
 0 Success                   8 Update applied
 1 Partial success           9 Update failed
 2 Failed                   10 Unsupported platform
 3 Bad arguments            11 Diagnostics only (not an error)
 4 Cancelled by user        12 Spooler service failure
 5 Not administrator        13 Restore failed
 6 Update check failed      14 Signature invalid
 7 Update available         15 Watchdog timeout
```

See [Exit Codes](Exit-Codes).

---

## Elevation

```cmd
Clear-PrintSpoolerQueue.bat /Elevate
```

The launcher re-launches itself through UAC via a temporary VBScript and `ShellExecute … "runas"`,
forwarding **the original arguments** (captured before the parse loop shifts them away) plus
`/NoPause`. The elevated window does the work; the original window reports and closes.

If the user declines the UAC prompt, exit code is **5** (`NotAdministrator`).

Without `/Elevate`, a non-elevated run fails immediately with exit 5 and instructions — it does not
half-run.

Admin detection uses three independent probes (`net session`, `fltmc`, `openfiles`), because
`net session` fails on hosts where the Server service is stopped even when the token *is* elevated.

---

## The "press any key" problem

A batch file that pauses under Task Scheduler blocks until the execution time limit kills the task.
This is exactly how the `schtasks` one-liner in the v3 README failed.

The launcher detects how it was started:

| Launched by | `CMDCMDLINE` looks like | Pauses? |
|---|---|---|
| Double-click | `cmd /c ""C:\path\file.bat" "` | Yes |
| Typed in a console | `"C:\path\file.bat" /Force` | Yes |
| Task Scheduler | `cmd /c C:\path\file.bat -Force` | **No** |

`/NoPause` always wins, and `PSQ_NONINTERACTIVE=1` forces non-interactive regardless.

Independently, the PowerShell side performs a stronger check (`[Console]::IsInputRedirected`,
`[Environment]::UserInteractive`) and refuses to prompt. Two layers, because a maintenance window
blocked on a Y/N question is a particularly annoying failure.

> **For scheduled tasks, do not use the `.bat` at all.** `Install-SpoolerCleanupTask` always launches
> the `.ps1` with `-Force -NonInteractive -Quiet -NoProgress`. See
> [Scheduling and Automation](Scheduling-and-Automation).

---

## Staged update swap

Before running, the launcher checks `<script dir>\.update\pending-update.json`. If a verified update
is waiting, it moves `Clear-PrintSpoolerQueue.ps1` into place so *this* run uses the new version.

It does **not** replace itself: `cmd.exe` re-reads a batch file by byte offset while it executes, so
overwriting it mid-run produces garbage. The launcher's own update lands on the following start, and
it says so:

```
[INFO] The launcher update will be applied on the next start.
```

The swap needs no hashing in batch, because those bytes were verified twice already — once at
download against `checksums.txt`, and again when written into `pending-update.json`. Keeping
PowerShell out of the launcher's hot path avoids the quoting hazards that make
`powershell -Command "…"` inside batch so fragile.

See [Updates and Versioning](Updates-and-Versioning).

---

## Encoding and portability constraints

The launcher is held to three rules, all asserted by the build:

1. **Pure ASCII.** Non-ASCII bytes render as mojibake in the OEM code page (437/850), and there is no
   reliable way to control that from inside a batch file.
2. **CRLF line endings.** An LF-only `.bat` misbehaves in `cmd.exe` — labels can fail to resolve.
3. **No tabs.** Tabs in batch produce unpredictable tokenisation.

`tools/Build.ps1` converts line endings and encodes the file as ASCII on every build, and the test
suite fails the build if any of the three is violated.

---

## Examples

```cmd
:: What is wrong? (read-only)
Clear-PrintSpoolerQueue.bat /Diagnostic

:: Fix the local queue with an archive
Clear-PrintSpoolerQueue.bat /Force /Backup

:: Only jobs older than two hours
Clear-PrintSpoolerQueue.bat /Force /Backup /OlderThan 2h

:: Deep clean a machine that keeps re-sticking
Clear-PrintSpoolerQueue.bat /Repair /Force /Backup /PurgeForms

:: Two servers, in parallel, on PowerShell 7
Clear-PrintSpoolerQueue.bat /Computers SRV01,SRV02 /Force /Backup /Throttle 8 /PS7

:: A fleet from a file, with a CSV for the record
Clear-PrintSpoolerQueue.bat /Hosts C:\Tools\hosts.txt /Force /Backup /ExportFormat Csv /ExportPath C:\Reports

:: Unattended, for a scheduler that insists on using the .bat
Clear-PrintSpoolerQueue.bat /Force /Quiet /NoPause /Backup /OlderThan 6h

:: Is there a newer version?
Clear-PrintSpoolerQueue.bat /CheckUpdate

:: Read the safety model without leaving the console
Clear-PrintSpoolerQueue.bat /Help Safety
```

---

## Launcher-only diagnostics

```cmd
Clear-PrintSpoolerQueue.bat /Version
```

prints the launcher version, the repository URL, and the `$Script:ScriptVersion` line read from the
adjacent `.ps1` — so you can see at a glance whether the two files are from the same build. A
mismatch means a partial update, which the staged-swap mechanism will resolve on the next start.

---

## Related pages

- [CLI Reference](CLI-Reference) — the PowerShell parameters these map to
- [Installation and Deployment](Installation-and-Deployment)
- [Scheduling and Automation](Scheduling-and-Automation)
- [Exit Codes](Exit-Codes)

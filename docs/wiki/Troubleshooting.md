# Troubleshooting

Symptom → cause → the exact command that fixes it.

Start here when something went wrong. If your problem is not listed, produce a diagnostics bundle
and [open an issue](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/issues) — see the last
section.

---

## Cleanup failures

### "The Spooler will not stop" / stuck in `STOP_PENDING`

**Cause.** Almost always `PrintNotify` ("Printer Extensions and Notifications") or a hung
`splwow64.exe` holding a handle. This is the single most common reason the classic
`net stop spooler` fails.

**Fix.**

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Repair -Force -Backup
```

`-Repair` stops `PrintWorkflowUserSvc_*`, `Fax` and `PrintNotify` **before** the Spooler. If it
still hangs, `Service.ForceStopFallback` (default `$true`) escalates to terminating the service
process — safe for the Spooler, because the SCM treats an exited service process as stopped.

**Confirm the diagnosis** by reading `serviceTrace[]` in the report, which records every transition
with its from/to state and duration:

```powershell
$run = .\Clear-PrintSpoolerQueue.ps1 -Repair -Force -Backup -PassThru
$run.Results[0].ServiceTrace | Format-Table Service, Action, From, To, Success, Note
```

---

### Files come straight back after the cleanup

**Cause.** A driver or port monitor is re-submitting them, or a client keeps resending. The report
flags this as `ResurrectedFiles > 0`.

**Fix.**

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic -ExportFormat Html -ExportPath C:\Temp
```

Look at the **Driver** and **Port** columns of the printer table, and at the driver list. Then:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Repair -PurgeDriverStore -Force -Backup
```

`-PurgeDriverStore` removes driver packages that **no** printer references — it cannot remove a
driver that is in use. If the offending driver is still attached to a printer, remove or replace
that printer's driver first.

---

### "Access is denied" on some files only

**Cause.** The file is held open by `spoolsv.exe`.

**Behaviour.** The tool retries with exponential backoff (500 ms → 1 s → 2 s) and reports the
remainder as `FilesLocked` with the exact per-file error. A locked file never aborts the run.

**Fix.** Usually none is needed — the exit code will be `1` (PartialSuccess) and the queue is
unblocked. To push harder:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Force -Backup -RetryCount 6 -RetryDelayMs 1000 -SettleSeconds 5
```

A longer `-SettleSeconds` gives the service more time to release handles after the stop, which is
the most effective knob.

---

### Exit code 2 with "Refusing to operate on…"

**Cause.** One of the four safety guards fired. The message names the guard.

| Message | Meaning |
|---|---|
| `Leaf 'X' must be 'PRINTERS'` | The resolved directory is not a spool directory |
| `must live under a 'spool' folder` | A corrupt `DefaultSpoolDirectory`, e.g. `C:\Users\PRINTERS` |
| `not rooted at a Windows drive letter` | A relative or POSIX-style path |
| `UNC spool directories are refused` | Clean the file server on the file server |
| `path is too shallow` | A truncated registry value |
| `equals a protected system root` | `C:\Windows` and friends |

**Fix.** Inspect the registry override:

```powershell
Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers' -Name DefaultSpoolDirectory
```

If it is genuinely corrupt, let the tool reset it:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Repair -ResetSpoolDirectory -Force
```

If the host legitimately uses a non-standard layout, pass it explicitly (still guarded):

```powershell
.\Clear-PrintSpoolerQueue.ps1 -SpoolDirectory 'D:\Print\spool\PRINTERS' -Force -Backup
```

**Never** work around a guard by editing the constant. The guard fired because something upstream is
wrong, and that something is worth knowing about.

---

### Exit code 2 with "…spool files found (guard limit 100000)"

**Cause.** The volume guard. Either a queue has genuinely exploded, or the directory resolved
somewhere unexpected.

**Fix.** `-Diagnostic` first. Check `SpoolDirectory.Path` and the per-extension breakdown. If the
path is wrong, this is the registry case above. If the path is right and there really are 100 000
files, you have an incident that needs a human, not a script.

---

## Remote execution

### All remote targets fail with a timeout

**Cause.** WinRM is not configured, or TCP 5985/5986 is blocked.

**Fix.**

```powershell
# On each target
Enable-PSRemoting -Force

# For workgroup (non-domain) machines, on BOTH sides
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "10.0.0.*" -Force

# Firewall
Get-NetFirewallRule -DisplayGroup "Windows Remote Management" | Enable-NetFirewallRule
```

Use `-PreFlightProbe` (on by default) to get the exact per-host reason in about a second each,
instead of a 40-second timeout per host:

```powershell
$run = .\Clear-PrintSpoolerQueue.ps1 -ComputerName (Get-Content .\hosts.txt) -Force -PassThru
$run.Results | Where-Object Status -eq 'Failed' | Format-Table ComputerName, Errors -Wrap
```

---

### Some hosts work, others do not

Mixed PowerShell versions are **fine** — the payload only uses cmdlets available in Windows
PowerShell 5.1, so it runs on an old Server 2012 R2 target and on PowerShell 7.6 alike.

Mixed *credentials* are not. Either the caller must be a local administrator on every target, or
you must pass `-Credential`:

```powershell
$cred = Get-Credential
.\Clear-PrintSpoolerQueue.ps1 -ComputerName $hosts -Credential $cred -Force -Backup
```

For HTTPS-only environments add `-UseSsl`.

---

### `-ThrottleLimit` seems to do nothing

`-ThrottleLimit > 1` uses `ForEach-Object -Parallel`, which needs **PowerShell 7**. On Windows
PowerShell 5.1 the tool automatically falls back to a runspace pool with the same throttling, so it
still parallelises — but check which engine you are on:

```powershell
$PSVersionTable.PSVersion
```

If you are being launched by the `.bat`, it prefers `pwsh.exe` when present; force it with `/PS7`.

---

## Selection filters

### `-PrinterName` matches nothing although jobs are visible

**Cause.** When the Spooler is hung it cannot answer WMI, so the inventory is built from `.SHD`
headers, where the printer name is best-effort. Jobs with no readable metadata cannot be matched by
name and are **deliberately not selected** — the run reports this in its warnings.

**Confirm.**

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic | Select-Object -ExpandProperty Results |
    Select-Object -ExpandProperty Jobs | Format-Table JobId, PrinterName, Source
```

`Source` is `WMI` (authoritative), `WMI+SHD` (correlated) or `SHD` (best-effort).

**Fix.** Filter by `-JobId` instead — the id is derived from the `.SHD` file name and is always
available — or use `-OlderThan`, which relies on file timestamps rather than metadata. Or drop the
filter when the whole queue must go.

---

### Everything was deleted although I passed a filter

Check the **execution plan** the tool prints before it asks for confirmation:

```
Selection       : jobs matching: printers like 'HP*'; older than 2h
```

If it says `ALL spool files (*.SPL, *.SHD)`, no filter reached the tool. The usual cause is a
parameter typo, or passing the filter positionally where it bound to `-ComputerName`.

`-WhatIf` shows the same plan with zero risk:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -PrinterName "HP*" -WhatIf
```

---

## Logging and output

### The log file is missing

The resolved log path is printed in the summary and recorded in the report. If the script lives
somewhere read-only, the tool falls back automatically:

`-LogPath` → `%ProgramData%\PrintSpoolerCleanup\Logs` → `%LocalAppData%\PrintSpoolerCleanup\Logs`
→ `<script dir>\Logs` → `%TEMP%\PrintSpoolerCleanup\Logs`

Force a location:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -LogPath C:\Admin\Logs\spooler.log -Force
```

A `Logging disabled: …` warning means every candidate failed — usually a locked-down machine where
even `%TEMP%` is redirected to a share.

---

### Garbled characters in cmd.exe output

The tool detects the console code page and switches to ASCII automatically when it cannot render
Unicode. Force it:

```cmd
Clear-PrintSpoolerQueue.bat /Ascii
```

or run `chcp 65001` first for full Unicode. Colour is also disabled automatically when output is
redirected, and when `NO_COLOR` is set.

---

### stdout is polluted and my RMM cannot parse it

Use `-MachineReadableJson`, which guarantees **exactly one JSON document on stdout and nothing
else** — no banner, no progress, no colour escapes. All human output goes to the log file.

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Force -Backup -MachineReadableJson | ConvertFrom-Json
```

Do not combine it with `-ExportFormat` (that combination is rejected with exit 3); use
`-ExportPath` for files and keep stdout for the document.

---

## Updates

### The update check always fails

**Cause.** GitHub is unreachable (proxy/firewall) or the anonymous rate limit (60 requests/hour per
source IP) is exhausted. Exit code **6** distinguishes this from "no update available" (0).

**Fix.**

```powershell
$env:HTTPS_PROXY = 'http://proxy.corp.local:8080'
$env:PSQ_GITHUB_TOKEN = 'ghp_...'        # 5000 req/h
.\Clear-PrintSpoolerQueue.ps1 -CheckUpdate
```

Air-gapped? Disable it permanently so it stops appearing in your logs:

```powershell
$env:PSQ_NO_UPDATE_CHECK = 'true'        # machine-wide via GPO preferences
```

---

### The update downloads but refuses to install

**Cause.** The release publishes no `checksums.txt`, and the tool will not install unverified bytes
by default. That is the intended behaviour.

**Fix (proper).** Publish `checksums.txt` with the release — `tools/Build.ps1` generates it
automatically.

**Fix (explicit override).**

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Update -Force -SkipChecksumVerification
```

---

### "The file is in use; the update stays staged"

Expected. A process cannot safely overwrite the file it is executing from, so updates land in
`.update\` and are swapped in at the **start of the next run**. Nothing is wrong; the new version
takes effect next time.

To force it now, close every instance and re-run, or pass `-ApplyUpdateNow`.

---

## Scheduling

### The task runs but does nothing

```powershell
Get-SpoolerCleanupTask | Format-List TaskName, State, LastRunTime, LastTaskResult, LastTaskResultMeaning, NextRunTime
```

| `LastTaskResult` | Meaning | Fix |
|---|---|---|
| `0x0` | Success | It worked; check the log for what it did |
| `0x1` | PartialSuccess | Some targets failed — read the report |
| `0x4` | Cancelled | **It was waiting for a confirmation nobody could give.** Add `-Force -NonInteractive` |
| `0x5` | NotAdministrator | Run the task as `SYSTEM` or an elevated account |
| `0x2` | Failed | Read the log; usually a guard or a service failure |

`0x4` is the classic hand-made-task failure. `Install-SpoolerCleanupTask` always adds
`-Force -NonInteractive -Quiet -NoProgress`, which is why you should use it rather than the raw
`schtasks` one-liner.

---

### The task blocks forever and is eventually killed

That is the v3 failure mode: the README recommended running the interactive `.bat` as SYSTEM, so
"Press any key" waited until the execution time limit fired.

The v4 installer never does this. If you built the task by hand, either add `/NoPause /Quiet` to
the `.bat` or — better — point the task at the `.ps1` with `-NonInteractive`.

---

## Printing still broken after a successful cleanup

Exit code `0` means the queue was cleared and the Spooler is running. It does not mean the printer
is reachable. Run `-Diagnostic` and check:

- **`WorkOffline`** on the printer — the queue is paused or the printer was marked offline
- **`PrinterStatus`** 3/4/5 — error, paper jam, or door open
- **Port `HostAddress`** — a TCP/IP port pointing at a dead or changed IP
- **`PrintNotify`** stopped — client-side printer extensions will not work
- **SCM 7031/7034 events** — the service crashed recently; schedule an `OnEvent` task to self-heal
- **`Source = SHD` on every job** — the Spooler was too sick to answer WMI

To prove printing really works rather than trusting the service state:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Force -Backup -SelfTestPrint
```

This submits a job through the real spool path and cancels it before it reaches the device.

---

## Getting help

Produce a bundle and attach it to an
[issue](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/issues):

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic -DiagnosticsBundlePath C:\Temp
```

The ZIP contains:

| File | Contents |
|---|---|
| `environment.json` | tool version and build, PowerShell edition and version, OS, elevation, console capability, culture, loaded modules |
| `run-report.json` / `.txt` | the full structured result |
| `log-buffer.txt` | the recent in-memory log ring buffer |
| `effective-config.json` | the configuration in effect, **with `Update.Token` redacted** |
| `file-integrity.json` | size, SHA-256 and Authenticode status of the installed files |
| `current.log` | a copy of the active log file |

It never contains credentials, GitHub tokens, or document contents. Please include:

1. The bundle
2. The exact command line you ran
3. The exit code
4. What you expected instead

Also worth running before you report:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -SelfTest       # is this copy of the tool internally sound?
.\Clear-PrintSpoolerQueue.ps1 -ShowVersion    # which version and build are you actually on?
```

# Scheduling and Automation

Automated maintenance is where this tool earns its keep: a wedged queue that self-clears at 03:00
never becomes a ticket.

---

## Install the task

```powershell
Install-SpoolerCleanupTask -ScriptPath C:\Tools\Clear-PrintSpoolerQueue.ps1 `
                           -Schedule Daily -StartTime 03:00 -PassThru
```

That is the whole installation. It registers
`PrintSpoolerCleanup\Clear Stuck Print Queue` running as `NT AUTHORITY\SYSTEM`.

| Parameter | Default | Notes |
|---|---|---|
| `-ScriptPath` | *(required)* | Full path to `Clear-PrintSpoolerQueue.ps1` |
| `-TaskName` | `PrintSpoolerCleanup\Clear Stuck Print Queue` | May include a folder |
| `-Schedule` | `Daily` | `Daily`, `Weekly`, `Hourly`, `AtStartup`, `OnEvent`, `Once` |
| `-StartTime` | `03:00` | Local time, `HH:mm` |
| `-IntervalHours` | `6` | For `-Schedule Hourly` |
| `-ComputerName` | local only | Targets the task should clean |
| `-ComputerListFile` | — | Better for large sites; the list can change without re-registering |
| `-Arguments` | — | Extra arguments verbatim, e.g. `'-OlderThan 6h -Repair'` |
| `-RunAsAccount` | `NT AUTHORITY\SYSTEM` | |
| `-ExecutionTimeLimitMinutes` | `30` | Task is terminated past this |
| `-RunOnlyIfIdle` | off | |
| `-NoBackup` / `-NoAutoUpdate` | off | Backup and auto-update are **on** by default |
| `-Method` | `Auto` | `Auto`, `Module` or `Xml` |

---

## What the task actually runs

```
<engine>.exe -NoProfile -ExecutionPolicy Bypass -File "<script>"
    -Force -NonInteractive -Quiet -NoProgress -Backup -AutoUpdate
```

Every one of those switches is there for a reason:

| Switch | Why |
|---|---|
| `-Force` | No confirmation prompt |
| `-NonInteractive` | Refuse to prompt rather than block forever |
| `-Quiet -NoProgress` | No console decoration nobody will read |
| `-Backup` | An unattended delete with no archive is how you lose someone's invoice |
| `-AutoUpdate` | Cooldown-limited, never blocks, never fails the run |

It always launches the **`.ps1`**, never the `.bat`. The `.bat` is a human-facing launcher with a
"press any key" pause; under Task Scheduler that pause blocks until the execution time limit kills
the task. This is precisely how the `schtasks` one-liner in the v3 README failed, and it is the
reason a real installer exists.

The engine is whichever PowerShell registered the task, so a PowerShell 7 deployment keeps
`-ThrottleLimit` support in the task.

---

## Schedules

```powershell
# Nightly at 03:00
Install-SpoolerCleanupTask -ScriptPath $s -Schedule Daily -StartTime 03:00

# Weekday mornings
Install-SpoolerCleanupTask -ScriptPath $s -Schedule Weekly -StartTime 06:30

# Every 6 hours
Install-SpoolerCleanupTask -ScriptPath $s -Schedule Hourly -IntervalHours 6

# Two minutes after boot — catches a spooler that died during an unclean shutdown
Install-SpoolerCleanupTask -ScriptPath $s -Schedule AtStartup

# Once, for a test
Install-SpoolerCleanupTask -ScriptPath $s -Schedule Once -StartTime 14:30
```

A random delay (default 15 minutes) is applied to calendar triggers so four hundred workstations do
not all hit your print servers in the same second.

---

## Event-driven self-healing

The highest-value schedule in the whole product. SCM event **7031** is "the service terminated
unexpectedly and was restarted" — exactly what a crashing Spooler produces:

```powershell
Install-SpoolerCleanupTask -ScriptPath C:\Tools\Clear-PrintSpoolerQueue.ps1 `
    -Schedule OnEvent -EventLogName System `
    -EventSource "Service Control Manager" -EventId 7031 `
    -Arguments '-OlderThan 1h'
```

The task fires within a minute of the crash, clears the queue that probably caused it, and restarts
cleanly. A driver that crashes the Spooler four times a night becomes four automatic recoveries
instead of four user tickets.

Other useful triggers:

| Event | Log / Provider | Meaning |
|---|---|---|
| `7031` | System / Service Control Manager | Service terminated unexpectedly, restarted |
| `7034` | System / Service Control Manager | Service terminated unexpectedly, **not** restarted |
| `7000` | System / Service Control Manager | Service failed to start |
| `7009` | System / Service Control Manager | Service start timed out |
| `315`, `316` | `Microsoft-Windows-PrintService/Admin` | Driver/rendering errors |

Omit `-EventId` to trigger on any event from that provider.

---

## Managing the task

```powershell
Get-SpoolerCleanupTask | Format-List *
Start-SpoolerCleanupTask -Wait              # run now and report the exit code
Uninstall-SpoolerCleanupTask
```

`Get-SpoolerCleanupTask` adds a `LastTaskResultMeaning` property that translates the numeric result
through the documented exit-code contract:

```
TaskName              : PrintSpoolerCleanup\Clear Stuck Print Queue
State                 : Ready
LastRunTime           : 10/10/2026 03:00:07
LastTaskResult        : 1
LastTaskResultMeaning : PartialSuccess
NextRunTime           : 11/10/2026 03:00:00
NumberOfMissedRuns    : 0
Action                : C:\Program Files\PowerShell\7\pwsh.exe -NoProfile -ExecutionPolicy Bypass ...
```

That last property is the point: `0x1` means *PartialSuccess* under v4, where under v3 it meant
*fatal*. Guessing from memory is how a healthy fleet gets escalated at 3 a.m.

---

## Settings the installer applies

These are the defaults that Task Scheduler's GUI would *not* give you, and each one corresponds to a
real failure:

| Setting | Value | Why |
|---|---|---|
| `DisallowStartIfOnBatteries` | **false** | Otherwise maintenance silently never runs on laptops |
| `StopIfGoingOnBatteries` | **false** | Same |
| `StartWhenAvailable` | **true** | A missed 03:00 run happens at the next opportunity |
| `ExecutionTimeLimit` | `PT30M` | A wedged run cannot consume the machine indefinitely |
| `RestartOnFailure` | 2 × 10 min | Transient WinRM failures recover on their own |
| `MultipleInstancesPolicy` | `IgnoreNew` | Two overlapping cleanups of the same queue is a bad idea |
| `RunLevel` | `HighestAvailable` | Required to manage the Spooler |
| `Priority` | 7 | Below interactive work |

`-ExecutionTimeLimitMinutes` raises the limit for large fleets. Pair it with the tool's own
`-Watchdog` when you want the *Spooler* recovered rather than merely the task terminated:

```powershell
Install-SpoolerCleanupTask -ScriptPath $s -Schedule Daily -StartTime 03:00 `
    -Arguments '-Watchdog -WatchdogMinutes 20' -ExecutionTimeLimitMinutes 30
```

The watchdog restarts the Spooler and terminates the run (exit 15) if it wedges, so a hang cannot
leave a host without printing. It runs out-of-process, so it survives the very hang it is guarding
against.

---

## Manual `schtasks` (for GPO or ConfigMgr packaging)

If you must package it without PowerShell:

```cmd
schtasks /Create /TN "PrintSpoolerCleanup" /SC DAILY /ST 03:00 /RU SYSTEM ^
  /TR "powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Tools\Clear-PrintSpoolerQueue.ps1 -Force -NonInteractive -Quiet -NoProgress -Backup" ^
  /RL HIGHEST /F
```

Note what is *not* there compared to the v3 README: no `.bat`, and `-Force -NonInteractive` are
present. Both differences are load-bearing.

---

## CI and RMM integration

For an agent that runs a command and reads stdout:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Force -Backup -OlderThan 6h `
    -NonInteractive -Quiet -NoProgress -MachineReadableJson
```

`-MachineReadableJson` guarantees **exactly one JSON document on stdout and nothing else** — no
banner, no progress records, no colour escape sequences. Everything human-readable goes to the log
file. The exit code carries the outcome; the document carries the detail.

Do not combine it with `-ExportFormat` (rejected with exit 3); use `-ExportPath` for files.

See [Reports and Machine-Readable Output](Reports-and-Machine-Readable-Output) for the schema.

---

## Group Policy deployment

Computer startup script, or a scheduled task pushed by GPO Preferences:

```
Computer Configuration → Policies → Windows Settings → Scripts → Startup
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File \\corp\share\Scripts\Clear-PrintSpoolerQueue.ps1 -Force -NonInteractive -Quiet -Backup -OlderThan 12h
```

Reading the script from a share means an update is a file copy. Combine with the tool's own update
mechanism and a share is not even required — see
[Updates and Versioning](Updates-and-Versioning).

For the task-based approach, deploy the XML:

```powershell
$def = Install-SpoolerCleanupTask -ScriptPath C:\Tools\Clear-PrintSpoolerQueue.ps1 -PassThru
$def.Xml | Set-Content \\corp\share\GPO\PrintSpoolerCleanup.xml -Encoding Unicode
```

Then distribute with GPO Preferences → Task Scheduler, which imports the same XML.

---

## Verifying it works

Do this once after installing, not six months later during an incident:

```powershell
# 1. Is the task registered and enabled?
Get-SpoolerCleanupTask | Format-List TaskName, State, NextRunTime

# 2. Does a manual run succeed?
Start-SpoolerCleanupTask -Wait
# -> Started : True
#    ExitCode: 0

# 3. Did it write a log?
Get-Content "$env:ProgramData\PrintSpoolerCleanup\Logs\PrintSpoolerCleanup.log" -Tail 20
```

A `LastTaskResult` of `0x4` after a manual run means the task is waiting for a confirmation that can
never arrive — it was created by hand and is missing `-NonInteractive`. Re-register it with
`Install-SpoolerCleanupTask`.

---

## Related pages

- [Exit Codes](Exit-Codes) — interpreting `LastTaskResult`
- [Remote and Fleet Management](Remote-and-Fleet-Management)
- [Updates and Versioning](Updates-and-Versioning)
- [Enterprise Deployment](Enterprise-Deployment)
- [Troubleshooting → Scheduling](Troubleshooting#scheduling)

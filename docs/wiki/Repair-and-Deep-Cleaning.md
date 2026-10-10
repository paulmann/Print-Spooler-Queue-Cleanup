# Repair and Deep Cleaning

`-Repair` is for the machine that keeps coming back. A normal cleanup clears the queue; a repair also
addresses the things that caused it to wedge in the first place.

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Repair -Force -Backup
```

`-Repair` performs the full cleanup sequence **plus** the repair steps below. It is a superset, not
an alternative — you do not need to run both.

---

## What a normal cleanup already does

Worth restating, because these are the parts that actually fix a hung stop and they happen in every
mode, not just `-Repair`:

1. Stop `PrintWorkflowUserSvc_*`, `Fax` and `PrintNotify` **before** the Spooler
2. Stop the Spooler, polling its real state rather than trusting the cmdlet's return code
3. Wait `-SettleSeconds` so `spoolsv.exe` releases its handles
4. Delete `.SPL`/`.SHD` with exponential-backoff retries
5. Restart the Spooler, then only the dependents that were actually running
6. Re-read the service state and re-scan the directory to confirm nothing refilled

---

## Repair steps

### `-PurgeFormsCache`

Clears the spooler Forms cache. A corrupt or malformed custom form — often imported from a driver
package or created by an application — can wedge the rendering path in a way that survives a queue
clear, because the form definition lives outside the spool directory.

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Repair -PurgeFormsCache -Force -Backup
```

The report tells you how much was cleared:

> *Repair: cleared 14 of 15 Forms cache file(s) in C:\Windows\System32\spool\PRINTERS\..\FORMS.*

Forms are rebuilt from the registry and from driver packages on demand, so clearing the cache is
safe. A file that cannot be deleted is counted and reported rather than aborting the step.

**Use when:** the same printer wedges repeatedly, `-Diagnostic` shows the queue emptying but jobs
still failing, or `PrintService/Admin` reports rendering errors (event 315/316).

---

### `-PurgeDriverStore`

Removes printer driver packages that **no printer references**.

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Repair -PurgeDriverStore -Force -Backup
```

This is the only destructive repair step, so it requires `-Force` as an explicit acknowledgement —
the tool rejects `-PurgeDriverStore` without `-Force` at parameter-validation time with exit code 3.

**What it cannot do:** remove a driver that is in use. The selection is
`Get-PrintDriver | Where-Object { -not (Get-Printer -DriverName $_.Name) }`, so any package still
attached to a printer is skipped. It also cannot fix a bad *active* driver — if the driver your
printers actually use is the problem, replace it; no cleanup tool can help.

**What it is genuinely good for:** print servers accumulate orphaned driver packages over years —
superseded versions, packages from printers that were deleted, duplicates from Windows Update
replacing a vendor driver with the Microsoft inbox one. These bloat the driver store, slow
`Add-Printer`, and are the thing a future deployment accidentally binds to.

> Report line: *Repair: removed 7 unreferenced printer driver package(s) from the driver store.*

**Use when:** after a driver-related crash loop (SCM 7031), before re-installing a vendor driver,
or as periodic hygiene on a long-lived print server.

---

### `-ResetSpoolDirectory`

Removes the `DefaultSpoolDirectory` registry override so the OS default is used again.

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Repair -ResetSpoolDirectory -Force
```

**Use when:** `-Diagnostic` shows `Spool directory` resolving somewhere unexpected, or the guards
refuse to operate because the registry value points at a path that is not a spool directory.

The tool already *validates* this value on every run — a UNC path, a value under four characters, or
one containing control characters is ignored with a warning and the OS default used instead.
`-ResetSpoolDirectory` goes further and removes the value outright, which is the right response when
you have confirmed it is simply wrong.

If the host **legitimately** uses a relocated spool directory (common on print servers, to keep the
system volume free), do not reset it — point the tool at it explicitly instead:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -SpoolDirectory 'D:\Print\spool\PRINTERS' -Force -Backup
```

---

## Dependent-service controls

These are on by default in every mode and are documented here because turning them *off* is
sometimes the right answer.

| Parameter | Default | Effect |
|---|---|---|
| `-StopDependentServices:$false` | `$true` | Leave `Fax`/`PrintNotify`/`PrintWorkflowUserSvc_*` alone |
| `-RestartDependentServices:$false` | `$true` | Do not restart them afterwards |
| `Service.ForceStopFallback` | `$true` | Permit terminating a service process stuck in `STOP_PENDING` |

**When to turn them off:** a host where `Fax` is genuinely in use and must not be interrupted, or an
environment where a management agent depends on `PrintNotify`. Turning them off makes the Spooler
stop more likely to hang — that trade-off is yours to make, and the report will show it.

`ForceStopFallback` should essentially never be disabled. It only fires after the normal stop has
already timed out, at which point the alternatives are "give up" and "terminate the wedged process".
For the Spooler, termination is safe: the SCM treats an exited service process as stopped and will
restart it according to its recovery policy.

---

## Verification: `-SelfTestPrint`

"Service state = Running" is not proof that printing works. This switch submits a job through the
real spool path and cancels it before it reaches the device:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Repair -Force -Backup -SelfTestPrint
```

> *Self-test print job on 'SRV01' was accepted and cancelled — the queue is functional.*

If submission fails, that is recorded as a warning on the result and the exit code reflects it —
which turns a silent "looks fixed" into a visible "not actually fixed".

If no printer is installed on the target, the check reports `NotApplicable` rather than failing.
Nothing is ever printed on a physical device.

---

## Recommended sequences

**A machine that keeps re-sticking, first pass:**

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic -ExportFormat Html -ExportPath C:\Temp
# read the report, then:
.\Clear-PrintSpoolerQueue.ps1 -Repair -PurgeFormsCache -Force -Backup -SelfTestPrint
```

**After a driver crash loop:**

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Repair -PurgeFormsCache -PurgeDriverStore -Force -Backup
Install-SpoolerCleanupTask -ScriptPath C:\Tools\Clear-PrintSpoolerQueue.ps1 -Schedule OnEvent `
    -EventLogName System -EventSource "Service Control Manager" -EventId 7031
```

**A print server with a corrupt spool path:**

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Repair -ResetSpoolDirectory -Force -Backup
```

**Periodic hygiene (monthly, off-hours):**

```powershell
.\Clear-PrintSpoolerQueue.ps1 -OlderThan 24h -Force -Backup -ExportFormat Csv -ExportPath \\filesrv\logs\spooler
```

Note this one is deliberately *not* `-Repair`: routine maintenance should clear old jobs, not purge
driver packages.

---

## What `-Repair` will not fix

Being straight about this saves a lot of time:

- **A bad active driver.** Replace it. `-PurgeDriverStore` only removes *unreferenced* packages.
- **A dead network printer or a wrong IP.** Fix the port. The queue will refill until you do.
- **A client that resubmits.** The report flags `ResurrectedFiles > 0`; find the submitter via
  `-OwnerName` / `UserName` in the job table.
- **Disk exhaustion.** Clearing helps immediately, but move the spool directory off the full volume.
- **A paused printer or "Use Printer Offline" on the client.** `-Diagnostic` shows
  `Offline = YES`; that is a printer-state problem, not a queue problem.

For each of these, [Why Print Queues Get Stuck](Why-Print-Queues-Get-Stuck) has the specific
workaround.

---

## Related pages

- [Diagnostics Guide](Diagnostics-Guide) — always diagnose before repairing
- [Safety Model](Safety-Model) — what the guards will and will not allow
- [Print Spooler Internals](Print-Spooler-Internals)
- [Backup and Restore](Backup-and-Restore)

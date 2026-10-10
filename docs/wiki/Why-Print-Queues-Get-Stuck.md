# Why Print Queues Get Stuck

Seven root causes account for almost every wedged queue. Each one presents differently, and knowing
which one you have determines whether clearing the queue is a fix or a temporary reprieve.

Run this first — it identifies the cause in about ten seconds and changes nothing:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic
```

---

## 1. A job that cannot be rendered

**The most common cause by a long way.**

A document is submitted, the driver fails to render it — a corrupt font, an unsupported page size,
an EMF record the driver chokes on, a 400-page job with transparency on a mono laserjet — and the
job neither completes nor fails cleanly. It sits in `Printing` or `Error` forever, and everything
queued behind it waits.

**How it presents.** Queue depth grows. The oldest job is hours or days old. One printer is affected
while others on the same server are fine.

**In `-Diagnostic`.** `Oldest file age >> newest file age`. A single job with an unusually large
`Size` or a `Status` of `Error`/`Printing` that never advances. `PrintService/Admin` events 315/316
(rendering failure).

**Fix.**

```powershell
.\Clear-PrintSpoolerQueue.ps1 -PrinterName "<that printer>" -Force -Backup
```

Clearing works. If the same document is resubmitted it will wedge again — tell the user, or the
ticket returns on Monday.

---

## 2. A dead or unreachable port

A TCP/IP port pointing at a printer whose IP changed (DHCP lease expired), or a WSD port whose
discovery entry went stale, or an SMB share that was removed.

**How it presents.** The printer shows `Offline` or `WorkOffline`. Jobs accumulate but nothing ever
prints. Often *all* jobs on that one printer.

**In `-Diagnostic`.** The printer table's `Offline` column says `YES`; the port table's `HostAddress`
does not resolve or does not answer.

**Fix.** Clearing the queue is only half the job — fix the port first, or the queue refills:

```powershell
Get-PrinterPort | Select-Object Name, HostAddress, PortNumber
Test-NetConnection <HostAddress> -Port 9100
```

Then clear:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -PrinterName "<that printer>" -Force -Backup
```

---

## 3. A dependent service holding the stop open

`PrintNotify`, `Fax` or `PrintWorkflowUserSvc_*` holds handles into the Spooler. The service will
not stop, so the classic `net stop spooler` fails and the operator cannot clear anything by hand.

**How it presents.** `net stop spooler` returns "The requested control is not valid for this
service" or simply hangs. `Get-Service Spooler` shows `StopPending`.

**In `-Diagnostic`.** `Spooler` state `Stop Pending`; `PrintNotify` running; `splwow64.exe` present
with a high thread count and `Responding = False`.

**Fix.** This is exactly what the tool exists for:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Repair -Force -Backup
```

Dependents are stopped first, then the Spooler. If it still hangs, `Service.ForceStopFallback`
(default on) terminates the wedged service process — safe, because the SCM treats an exited service
process as stopped.

See [Print Spooler Internals](Print-Spooler-Internals#why-net-stop-spooler-hangs).

---

## 4. A faulty or incompatible driver

A driver that crashes, leaks, or blocks. Common after a Windows update silently replaced a vendor
driver with the Microsoft inbox one, or when a 32-bit driver meets a 64-bit host.

**How it presents.** The Spooler crashes repeatedly and restarts (SCM 7031 events). Clearing the
queue works for a while, then the same thing happens. `spoolsv.exe` working set grows between
restarts.

**In `-Diagnostic`.** `EventLog` shows repeated 7031/7034. `SpoolerProcess` shows a recent
`StartTime` and a large `WorkingSetMB`. The driver list shows a driver whose `Manufacturer` is
`Microsoft` where you expected a vendor name.

**Fix.**

```powershell
# Deep clean, including the Forms cache and unreferenced driver packages
.\Clear-PrintSpoolerQueue.ps1 -Repair -PurgeFormsCache -PurgeDriverStore -Force -Backup
```

`-PurgeDriverStore` only removes packages that **no** printer references, so it cannot break a
working printer — but it does clear out the orphaned package that a future `Add-Printer` might pick
up again. If the *active* driver is the problem, replace it; no cleanup tool can fix a bad driver.

Schedule self-healing for the crashes:

```powershell
Install-SpoolerCleanupTask -ScriptPath C:\Tools\Clear-PrintSpoolerQueue.ps1 -Schedule OnEvent `
    -EventLogName System -EventSource "Service Control Manager" -EventId 7031
```

---

## 5. A corrupt spool directory or registry override

`DefaultSpoolDirectory` points somewhere wrong — a path that no longer exists, a UNC path to a
share that is offline, or a truncated value from a bad GPO edit. Or the directory's ACL was
modified and `spoolsv.exe` can no longer write.

**How it presents.** Jobs vanish on submit, or the Spooler starts and immediately stops, or
printing works for some users and not others.

**In `-Diagnostic`.** `SpoolDirectory.Path` is not what you expected; `AclOwner` is not
`NT AUTHORITY\SYSTEM`; `FileCount` is 0 while users report jobs disappearing.

**Fix.**

```powershell
Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers' -Name DefaultSpoolDirectory

# Reset to the OS default
.\Clear-PrintSpoolerQueue.ps1 -Repair -ResetSpoolDirectory -Force
```

> The tool validates this registry value rather than trusting it. A UNC path, a value shorter than
> four characters, or one containing control characters is ignored with a warning and the OS default
> is used — because acting on a corrupt value would point a recursive delete at a file share.

---

## 6. Disk exhaustion

The spool volume is full. New jobs cannot be written; existing ones cannot be read back.

**How it presents.** Everything fails at once, on every printer. Event log full of disk errors.

**In `-Diagnostic`.** `TotalBytes` in the spool directory is enormous; `Free memory` is fine but the
volume is not reported as healthy.

**Fix.** Clearing is exactly right here, and `-Backup` should be pointed at a *different* volume:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Force -Backup -BackupPath D:\SpoolBackups
```

Then move the spool directory off the system volume, which is good practice on any busy print server
regardless.

---

## 7. A client that will not stop resubmitting

The job is cleared, and immediately reappears. A print client retries, a workflow system polls and
resubmits, or a user's document is still open in the application and being re-sent.

**How it presents.** The cleanup succeeds — exit code 0 — but the queue is not empty five minutes
later.

**In the report.** `ResurrectedFiles > 0` with the warning:

> *N spool file(s) reappeared after the service restart. A driver or port monitor is re-submitting
> them; run -Diagnostic and check the driver/port.*

**Fix.** Find the submitter before clearing again, or you will be here every hour:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic | Select-Object -ExpandProperty Results |
    Select-Object -ExpandProperty Jobs | Format-Table JobId, PrinterName, UserName, DocumentName, Submitted
```

`UserName` and `DocumentName` tell you who and what. Then either pause the printer, or clear with
the queue held:

```powershell
Set-Printer -Name "<printer>" -PrinterStatus Paused
.\Clear-PrintSpoolerQueue.ps1 -PrinterName "<printer>" -Force -Backup
Set-Printer -Name "<printer>" -PrinterStatus Normal
```

---

## Not a queue problem at all

Sometimes the Spooler is perfectly healthy and users still cannot print. `-Diagnostic` will show a
clean queue and the real cause elsewhere:

| Symptom in the report | Actual problem |
|---|---|
| Printer `Offline = YES` | Paused queue, or "Use Printer Offline" set on the client |
| `PrintNotify` stopped and start type Manual | Client printer extensions and notifications unavailable |
| Port `HostAddress` unreachable | Network or the printer itself |
| Driver `Manufacturer = Microsoft` on a vendor device | Windows Update replaced the vendor driver |
| Zero printers listed | The Print Services role is not installed, or the printer was deleted |

Clearing the queue in these cases makes the ticket look closed for an hour and then it comes back.
Read the diagnostics first.

---

## The decision tree

```
Users cannot print
        |
        +-- Is the queue actually stuck?  ->  -Diagnostic
        |        |
        |        +-- Queue depth 0  ->  NOT a queue problem (port / offline / driver)
        |        +-- Queue depth > 0, oldest job hours old  ->  cause 1, 2 or 7
        |        +-- Spooler state Stop Pending  ->  cause 3  ->  -Repair
        |        +-- Repeated SCM 7031 events  ->  cause 4  ->  -Repair -PurgeDriverStore
        |        +-- Spool directory empty but jobs vanish  ->  cause 5 or 6
        |
        +-- Can the service be stopped?
                 |
                 +-- Yes  ->  -Force -Backup  (add -OlderThan on a busy shared queue)
                 +-- No   ->  -Repair -Force -Backup
```

---

## Related pages

- [Print Spooler Internals](Print-Spooler-Internals) — the mechanisms behind these failures
- [Diagnostics Guide](Diagnostics-Guide) — reading the output in detail
- [Repair and Deep Cleaning](Repair-and-Deep-Cleaning)
- [Troubleshooting](Troubleshooting) — when the tool itself misbehaves

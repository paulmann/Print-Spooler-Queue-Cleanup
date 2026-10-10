# Remote and Fleet Management

One invocation can service a single workstation or four hundred of them. The same payload runs on
every target, so remote behaviour is identical to local behaviour by construction rather than by
convention.

```powershell
.\Clear-PrintSpoolerQueue.ps1 -ComputerName SRV-PRINT01, SRV-PRINT02 -Force -Backup
```

---

## Transport: WinRM over `Invoke-Command`

This is the only transport that behaves identically on Windows PowerShell 5.1 and PowerShell 7.
The `-ComputerName` parameter of `Get-Service` and `Stop-Service` was **removed in PowerShell 6**,
which is exactly the bug v3.0 shipped with: a script that worked on 5.1 and threw on 7.

Consequence: targets must have WinRM enabled and reachable. There is no SMB-only or WMI-only mode.

### Prerequisites on each target

```powershell
# 1. Enable WinRM and the firewall rules
Enable-PSRemoting -Force

# 2. Confirm the listener
winrm enumerate winrm/config/listener

# 3. Firewall (usually handled by Enable-PSRemoting)
Get-NetFirewallRule -DisplayGroup "Windows Remote Management" | Enable-NetFirewallRule
```

Ports: **TCP 5985** (HTTP) or **TCP 5986** (HTTPS, with `-UseSsl`).

### Authorization

The caller must be a local administrator on the target, **or** you pass `-Credential`:

```powershell
$cred = Get-Credential
.\Clear-PrintSpoolerQueue.ps1 -ComputerName $hosts -Credential $cred -Force -Backup
```

Credentials live in a `PSCredential` for the life of the process only. They are never written to the
log file, the report, or the diagnostics bundle, and they are never forwarded to an elevated child
process.

### Workgroup (non-domain) machines

Kerberos is not available, so both sides need `TrustedHosts`:

```powershell
# On the machine running the tool
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "10.0.0.*" -Force

# Or use HTTPS with a certificate the caller trusts
.\Clear-PrintSpoolerQueue.ps1 -ComputerName 10.0.0.15 -UseSsl -Credential $cred -Force
```

In a domain, Kerberos handles this and no `TrustedHosts` entry is needed.

---

## Pre-flight probing

`-PreFlightProbe` is **on by default**. Each remote target gets a trivial `Invoke-Command` with a
short timeout before any real work starts.

Without it, every offline host costs a full WinRM timeout — roughly 40 seconds — *before* failing.
Across 400 hosts with 30 offline, that is twenty minutes of pure waiting, and because the failures
arrive late they also arrive interleaved with real results.

With it, offline hosts are identified in about a second each and reported with the specific reason:

```
[WARN] [SRV-042] WinRM probe failed: No WinRM response within 10s (host offline, WinRM
       disabled, or a firewall blocking 5985/5986).
```

The probe runs inside a job, so a *hung* WinRM stack cannot block the caller beyond
`Remote.ProbeTimeoutSeconds` (default 10).

Disable it only when you have already verified reachability by other means:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -ComputerName $hosts -PreFlightProbe:$false -Force
```

---

## Parallelism

```powershell
.\Clear-PrintSpoolerQueue.ps1 -ComputerName (Get-Content .\hosts.txt) -ThrottleLimit 24 -Force -Backup
```

| Engine | Mechanism |
|---|---|
| PowerShell 7.x | `ForEach-Object -Parallel -ThrottleLimit N` |
| Windows PowerShell 5.1 | `System.Management.Automation.Runspaces` pool with the same throttling |

The fallback is automatic — the same command line works on both, so you do not need to know which
engine a given technician is running.

**The default is 1 (sequential), and that is deliberate.** A spooler restart interrupts printing on
that host for several seconds. Running 64 at once across a site during business hours is a
self-inflicted outage that happens to look like a maintenance task. Raise the throttle knowingly,
off-hours, or against hosts you have already grouped by location.

Sizing guidance:

| Fleet | Throttle | Notes |
|---|---|---|
| < 20 hosts | 1 | Sequential is fast enough and simplest to reason about |
| 20–200 | 8–16 | Bounded by your network and the source host's CPU |
| 200–1000 | 16–32 | Watch the source host's memory; each worker holds a session |
| > 1000 | Chunk it | Split by AD site or subnet and run several invocations |

Maximum is 64 (`MaxThrottleLimit`).

### Isolation guarantee

**One failing computer can never abort the fleet.** Each worker is isolated, every target produces
exactly one result record, and exceptions are captured into that record rather than propagated. A
host that is offline, unauthorized, or whose Spooler refuses to stop shows up as a `Failed` row
alongside the successes.

---

## Supplying the target list

Three ways, and they merge (de-duplicated):

```powershell
# 1. Explicit
-ComputerName SRV01, SRV02, SRV03

# 2. Pipeline — binds by value or by property name
Get-ADComputer -Filter {OperatingSystem -like "Windows*"} -SearchBase "OU=Print,DC=corp,DC=local" |
    Clear-PrintSpoolerQueue -Force -Backup

Import-Csv .\assets.csv | Clear-PrintSpoolerQueue -Force        # needs a ComputerName column

# 3. File — better for large sites
-ComputerListFile C:\Tools\hosts.txt
```

`hosts.txt` format: one name per line; blank lines and lines starting with `#` or `;` are ignored.

```
# Print servers - site A
SRV-PRINT01
SRV-PRINT02

# Workstations - finance floor
PC-FIN-01
PC-FIN-02
```

A file is preferable to a hard-coded list for a scheduled task: the list can be updated by whoever
owns the asset inventory without touching the task definition, and without re-registering anything.

There is also `PSQ_COMPUTERS` for machine-wide policy.

---

## Mixed environments

**Mixed PowerShell versions are fine.** The payload uses only cmdlets available in Windows PowerShell
5.1, so the same code runs against Server 2012 R2 and against a Windows 11 host with PowerShell 7.6.
This is a hard constraint on the payload, asserted by review rather than by hope.

**Mixed credentials are not.** Either the caller is a local administrator everywhere, or you pass
`-Credential` that works everywhere. There is no per-host credential map — that would put a
credential table in a script.

**Mixed architectures are fine** (x86/x64 targets). The payload is architecture-neutral.

---

## Reading fleet results

```powershell
$run = .\Clear-PrintSpoolerQueue.ps1 -ComputerName $hosts -ThrottleLimit 16 -Force -Backup -PassThru

# What failed, and why
$run.Results | Where-Object Status -eq 'Failed' |
    Select-Object ComputerName, @{ n = 'Reason'; e { $_.Errors -join '; ' } } |
    Format-Table -Wrap

# Which hosts are left without a running spooler (the urgent ones)
$run.Results | Where-Object { $_.ServiceIsRunning -eq $false } |
    Select-Object ComputerName, Status

# Where files survived the delete
$run.Results | Where-Object { $_.FilesLocked -gt 0 -or $_.ResurrectedFiles -gt 0 } |
    Select-Object ComputerName, FilesLocked, ResurrectedFiles

# Every backup created, with its hash
$run.Results | Where-Object BackupPath | Select-Object ComputerName, BackupPath, BackupHash
```

For a durable record:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -ComputerName $hosts -ThrottleLimit 16 -Force -Backup `
    -ExportFormat Csv, Html -ExportPath \\filesrv\logs\spooler
```

The CSV is one row per host, which is what you want in a spreadsheet or a SIEM. See
[Reports and Machine-Readable Output](Reports-and-Machine-Readable-Output).

---

## Health probing without cleaning

`Test-PrintSpoolerHealth` is the cheap, non-destructive version — a 0–100 score per host, suitable
for a recurring dashboard rather than an incident:

```powershell
Get-Content .\hosts.txt | Test-PrintSpoolerHealth -ThrottleLimit 16 |
    Where-Object Health -ne 'Healthy' |
    Format-Table ComputerName, Health, Score, QueueDepth, OldestJobHours, Findings -Wrap
```

`Health` is `Healthy` (≥85), `Degraded` (≥60), `Unhealthy` (<60) or `Unknown` (unreachable). It never
stops a service and never deletes a file.

---

## Performance

Per host, the floor is:

```
stop dependents (~1–3 s) + stop spooler (~1–5 s) + settle (2 s default)
  + delete (proportional to file count and lock contention)
  + start spooler (~2–4 s) + start dependents (~1–3 s) + verify (~2 s)
```

Roughly **10–20 seconds per host** for a typical queue, dominated by fixed service-transition cost
rather than by file count. Sequential across 400 hosts is therefore 1–2 hours; at
`-ThrottleLimit 24` it is under ten minutes.

Reduce the fixed cost only when you understand the trade-off: `-SettleSeconds 0` saves two seconds
per host and markedly increases locked-file failures.

---

## Troubleshooting remote runs

See [Troubleshooting → Remote execution](Troubleshooting#remote-execution) for the specific cases.
The short version:

1. `-PreFlightProbe` gives you the per-host reason fast — read the warnings first.
2. Test one host interactively before running the fleet.
3. `Test-WSMan -ComputerName <host>` isolates WinRM from the tool entirely.
4. A host that is fine interactively but fails from a scheduled task is a credential problem, not a
   network problem.

---

## Related pages

- [Scheduling and Automation](Scheduling-and-Automation)
- [Reports and Machine-Readable Output](Reports-and-Machine-Readable-Output)
- [Exit Codes](Exit-Codes)
- [Enterprise Deployment](Enterprise-Deployment)

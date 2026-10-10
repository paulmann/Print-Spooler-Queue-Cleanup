# Enterprise Deployment

Rolling this out across a site is mostly a policy exercise: where the files live, what the defaults
are, who can run it, where the evidence goes, and how you know it worked.

---

## Reference architecture

```
                    ┌──────────────────────────────────────┐
                    │  \\corp-fs01\share\Scripts\PSQ\      │
                    │    Clear-PrintSpoolerQueue.ps1       │
                    │    Clear-PrintSpoolerQueue.bat       │
                    │    PrintSpoolerCleanup.config.psd1   │  ← site defaults
                    └──────────────────┬───────────────────┘
                                       │ read at run time
        ┌──────────────────────────────┼──────────────────────────────┐
        │                              │                              │
┌───────▼────────┐            ┌────────▼────────┐            ┌────────▼────────┐
│  Helpdesk      │            │  Scheduled task │            │  RMM agent      │
│  ad-hoc runs   │            │  (SYSTEM, 03:00)│            │  JSON + exit    │
└───────┬────────┘            └────────┬────────┘            └────────┬────────┘
        │                              │                              │
        └──────────────────────────────┼──────────────────────────────┘
                                       │
                    ┌──────────────────▼───────────────────┐
                    │  Evidence sinks                      │
                    │    logs    → %ProgramData% or share  │
                    │    backups → \\corp-fs01\SpoolerBk   │
                    │    reports → \\corp-fs01\Logs\Spooler│
                    │    events  → Application log / SIEM  │
                    └──────────────────────────────────────┘
```

Three consumption paths, one copy of the tool, one configuration file, one set of evidence sinks.

---

## Deployment steps

### 1. Choose the distribution point

A read-only share is fine — the tool never needs to write next to itself. It resolves its log target
by probing `%ProgramData%`, `%LocalAppData%`, the script directory and `%TEMP%` in order, and uses
the first one it can actually write to.

### 2. Write the site configuration

`PrintSpoolerCleanup.config.psd1` next to the script is auto-discovered by every run.

```powershell
@{
    Cleanup = @{
        Backup          = $true
        BackupPath      = '\\corp-fs01\share\SpoolerBackups'
        BackupKeepDays  = 30
        OlderThan       = '6h'          # protect jobs that are legitimately printing
        RetryCount      = 4
    }
    Service = @{
        TimeoutSeconds  = 120           # large queues on slow storage
        SettleSeconds   = 3
    }
    Logging = @{
        Level           = 'INFO'
        RetentionDays   = 90
        MaxSizeMB       = 20
        EventLog        = $true         # one SIEM query gives you the whole site
    }
    Execution = @{
        ThrottleLimit   = 16
        ExportFormat    = @('Csv')
        ExportPath      = '\\corp-fs01\share\Logs\Spooler'
        NonInteractive  = $true
        WatchdogMinutes = 20
    }
    Update = @{
        Channel         = 'stable'
        CooldownHours   = 168
        VerifyChecksum  = $true
    }
}
```

Two of these deserve emphasis:

- **`OlderThan = '6h'` as a site default** is the single highest-value policy setting. It means an
  automated run can never destroy a job that was submitted five minutes ago. A technician who needs
  the whole queue gone passes no filter and gets a prompt.
- **`EventLog = $true`** gives you one `Get-WinEvent` query for the entire site's maintenance history
  without shipping files anywhere.

Validation is strict: an unknown key or an out-of-range value aborts with exit 3 rather than being
ignored. Test the file before publishing it:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -ConfigPath .\PrintSpoolerCleanup.config.psd1 -WhatIf
```

### 3. Deploy the files

GPO startup script, ConfigMgr package, Intune, or a plain `robocopy`. Then verify:

```powershell
& '\\corp-fs01\share\Scripts\PSQ\Clear-PrintSpoolerQueue.ps1' -SelfTest -Quiet
if ($LASTEXITCODE -ne 0) { throw 'Self-test failed' }
```

### 4. Register the scheduled task

```powershell
Install-SpoolerCleanupTask -ScriptPath '\\corp-fs01\share\Scripts\PSQ\Clear-PrintSpoolerQueue.ps1' `
    -Schedule Daily -StartTime 03:00 -ComputerListFile '\\corp-fs01\share\Scripts\PSQ\print-hosts.txt' `
    -PassThru
```

One task on a management server can service the whole site via `-ComputerListFile`. Alternatively
register a local task on each host with no `-ComputerName`, which is more resilient (no single
management host) but harder to update centrally.

| Approach | Choose when |
|---|---|
| Central task + computer list | You have a management server and want one place to change |
| Local task per host | Hosts are often off-network, or you have no management tier |
| RMM invocation | You already have an agent; use `-MachineReadableJson` |

### 5. Add the self-healing trigger

```powershell
Install-SpoolerCleanupTask -TaskName 'PrintSpoolerCleanup\Self-Heal After Crash' `
    -ScriptPath '\\corp-fs01\share\Scripts\PSQ\Clear-PrintSpoolerQueue.ps1' `
    -Schedule OnEvent -EventLogName System -EventSource 'Service Control Manager' -EventId 7031 `
    -Arguments '-OlderThan 1h'
```

A driver that crashes the Spooler four times a night becomes four automatic recoveries instead of
four tickets.

### 6. Verify, then measure

```powershell
Start-SpoolerCleanupTask -Wait                      # ExitCode should be 0 or 1
Get-SpoolerCleanupTask | Format-List State, LastRunTime, LastTaskResult, LastTaskResultMeaning
Import-Csv '\\corp-fs01\share\Logs\Spooler\*.csv' | Measure-Object -Property FilesDeleted -Sum
```

---

## Who may run it

The tool requires **Administrator** for the local machine, because managing the Spooler and deleting
files in `PRINTERS` both need it. There is no reduced-privilege mode, and inventing one would be
worse than not having it: a partial cleanup that cannot stop the service leaves the queue in a worse
state than it started.

| Role | Capability |
|---|---|
| Local Administrator | Everything, locally |
| Domain admin / delegated local admin on print servers | Fleet-wide via WinRM |
| `NT AUTHORITY\SYSTEM` (scheduled task) | Everything on the host it runs on |
| Standard user | Nothing. Exit code 5 |

For helpdesk staff who are not local administrators, run the task centrally as SYSTEM rather than
granting admin rights — a scheduled task with a fixed, reviewed argument list is a far smaller grant
than local admin.

`-SelfElevate` will prompt through UAC for an interactive user who *can* elevate:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -SelfElevate -Force -Backup
```

---

## Evidence and audit

Four independent records, each answering a different question:

| Question | Source |
|---|---|
| "What did the tool do?" | The **log file** — every action, timestamped UTC, pipe-delimited or JSON Lines |
| "What was the outcome?" | The **report** (JSON/CSV/HTML/Text) — per-host status, counters, warnings, errors |
| "Did anything get recorded centrally?" | The **Event Log** — source `PrintSpoolerCleanup`, IDs 1000–1003 |
| "What exactly was deleted?" | The **backup archive** — the files plus a manifest listing job metadata and the archive SHA-256 |

Every report records `Operator`, `HostComputer`, `PowerShellVersion`, `ConfigSource`, the selection
filters in force, and `ExitCode`. So a run is attributable and reproducible: you can tell who ran it,
from where, with which configuration, against what, under which filters.

Retention is governed by `Logging.RetentionDays` and `Cleanup.BackupKeepDays`. For a regulated
environment set both high and point them at a WORM share:

```powershell
@{
    Logging = @{ RetentionDays = 365; EventLog = $true }
    Cleanup = @{ Backup = $true; BackupRequired = $true; BackupKeepDays = 365
                 BackupPath = '\\audit-share\SpoolerBackups' }
}
```

`-BackupRequired` makes a failed backup abort the run — the right trade-off where the queue has legal
significance.

---

## Monitoring

### Health (non-destructive, poll frequently)

```powershell
Get-Content .\print-hosts.txt | Test-PrintSpoolerHealth -ThrottleLimit 16 |
    Select-Object ComputerName, Health, Score, QueueDepth, OldestJobHours, Findings |
    ConvertTo-Json -Depth 4
```

Alert on `Health -ne 'Healthy'`. `Findings` carries the reason, which is what makes the alert
actionable rather than noise.

### Maintenance outcome (event-driven)

```powershell
Get-WinEvent -FilterHashtable @{ LogName='Application'; ProviderName='PrintSpoolerCleanup'; StartTime=(Get-Date).AddHours(-24) } |
    Select-Object TimeCreated, Id, LevelDisplayName, Message
```

ID 1002 (Error) and 1001 (Warning) are the ones to surface.

### Exit-code alerting

| Alert on | Codes | Meaning |
|---|---|---|
| **Page** | 2, 5, 12, 14, 15 | Nothing cleaned / not elevated / service failure / bad signature / watchdog |
| **Ticket** | 1, 9, 13 | Partial success, update failed, restore failed |
| **Dashboard only** | 6, 7 | Update check failed, update available |
| **Healthy** | 0, 8, 11 | Success, updated, diagnostics only |

Full table: [Exit Codes](Exit-Codes).

### Trending

The CSV export is one row per host per run. Load it anywhere and you can answer:

- which print servers wedge most often
- which drivers correlate with `FilesLocked` and `ResurrectedFiles`
- how much spool volume each site generates
- whether mean time to clear has improved

---

## Rollout plan

A staged rollout costs an afternoon and saves a week.

**Phase 1 — observe only (week 1).** Deploy the files and the config. Run `-Diagnostic` on a pilot
group on a schedule, exporting CSV. No cleanup yet. You now have a baseline: how many hosts have
wedged queues, how deep, how old.

```powershell
.\Clear-PrintSpoolerQueue.ps1 -ComputerName $pilot -Diagnostic -ExportFormat Csv `
    -ExportPath \\corp-fs01\share\Logs\Spooler-baseline
```

**Phase 2 — conservative cleanup (week 2).** Enable the daily task with a wide `OlderThan` (24 h) and
`-Backup`. Watch the reports. Nothing should be surprising; if it is, you learned that cheaply.

**Phase 3 — tighten (week 3+).** Reduce `OlderThan` to 6 h, add the OnEvent self-healing task, extend
to the full fleet in AD-site-sized chunks.

**Phase 4 — measure.** Compare the phase-1 baseline against steady state. Print-related ticket volume
is the number that matters, and it is the one that justifies the whole exercise.

---

## Air-gapped and restricted environments

| Constraint | Setting |
|---|---|
| No internet | `Update = @{ Enabled = $false; AutoCheck = $false }` or `PSQ_NO_UPDATE_CHECK=true` |
| Proxy required | `$env:HTTPS_PROXY`; the updater uses `Invoke-RestMethod`, which honours it |
| GitHub API rate-limited | `PSQ_GITHUB_TOKEN` (5000 req/h vs 60) |
| Internal mirror | `Update.Repository` and `Update.Branch`; mirror `checksums.txt` with the files |
| `AllSigned` execution policy | Sign with `tools/Build.ps1 -SigningThumbprint …` and run with `-VerifySignature` |
| No admin rights for staff | Central scheduled task as SYSTEM |
| Read-only distribution share | Supported; logs fall back to `%ProgramData%` / `%LocalAppData%` / `%TEMP%` |
| No Event Log write access | Event logging disables itself for the run and logs a warning; the file log continues |

---

## Common mistakes

| Mistake | Consequence | Instead |
|---|---|---|
| Scheduling the `.bat` | "Press any key" blocks until the time limit kills the task | `Install-SpoolerCleanupTask`, which always uses the `.ps1` with `-NonInteractive` |
| A hand-made task without `-Force` | Exit 4 every night; nothing is ever cleaned | Use the installer |
| No `-Backup` on a site with legal print traffic | An unrecoverable delete | `Cleanup.Backup = $true`, and `-BackupRequired` where it matters |
| No `OlderThan` default | An automated run destroys a job that was printing fine | `Cleanup.OlderThan = '6h'` site-wide |
| Editing the config by hand on each host | Drift, and nobody knows the real defaults | One config file on the share |
| Alerting on every non-zero exit code | Alert fatigue; the real failures get missed | Page on 2/5/12/14/15 only |
| Editing `dist/` instead of `src/` | The next build overwrites your change | Edit `src/`, run `tools/Build.ps1` |

---

## Related pages

- [Installation and Deployment](Installation-and-Deployment)
- [Configuration](Configuration)
- [Scheduling and Automation](Scheduling-and-Automation)
- [Remote and Fleet Management](Remote-and-Fleet-Management)
- [Security and Privacy](Security-and-Privacy)
- [Reports and Machine-Readable Output](Reports-and-Machine-Readable-Output)

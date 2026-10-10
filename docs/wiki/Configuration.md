# Configuration

Everything the tool does can be defaulted once, per site, instead of being typed on every command
line. v3 had no configuration story at all: changing the settle delay or the log location meant
editing the script, which broke the next update.

---

## Precedence

Highest wins:

| # | Source | Example |
|---|---|---|
| 1 | Command-line parameters | `-ServiceTimeoutSeconds 120` |
| 2 | `PSQ_*` environment variables | `PSQ_SERVICE_TIMEOUT=120` |
| 3 | `-ConfigPath <file>` | `-ConfigPath C:\Tools\site.psd1` |
| 4 | Auto-discovered config file | `PrintSpoolerCleanup.config.psd1` |
| 5 | Built-in defaults | `Service.TimeoutSeconds = 60` |

The run report and log record which source was used (`ConfigSource`), so you can always tell why a
value was what it was:

```
Configuration source: AutoDiscover+Environment (C:\Tools\PrintSpoolerCleanup.config.psd1)
```

### Auto-discovery order

1. `<script directory>\PrintSpoolerCleanup.config.psd1`
2. `<current directory>\PrintSpoolerCleanup.config.psd1`
3. `%ProgramData%\PrintSpoolerCleanup\PrintSpoolerCleanup.config.psd1` ← **machine-wide policy**
4. `%LocalAppData%\PrintSpoolerCleanup\PrintSpoolerCleanup.config.psd1` ← per-user

The first file found wins. `%ProgramData%` is the right place for a site-wide default pushed by GPO.

`-NoConfigFile` ignores auto-discovery entirely, which is useful when you need a reproducible run
regardless of what is installed on the host.

---

## Generating a starting point

The shipped template is fully commented, but you can also emit the effective configuration — which
is the defaults merged with anything already in force:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -SaveConfig -ConfigPath C:\Tools\PrintSpoolerCleanup.config.psd1
```

---

## The schema

### `Service`

| Key | Default | Range | Notes |
|---|---|---|---|
| `Name` | `'Spooler'` | | Only change this on an unusual SKU |
| `TimeoutSeconds` | `60` | 5–3600 | Seconds to wait for a state change. A large queue on slow storage needs more than the classic 30 |
| `SettleSeconds` | `2` | 0–300 | Pause after the stop so handles are released. `0` markedly increases locked-file failures |
| `StopDependentServices` | `$true` | | Stop `Fax`/`PrintNotify`/`PrintWorkflowUserSvc_*` first — the setting that most often turns a hung stop into a clean run |
| `RestartDependentServices` | `$true` | | Only services that were `Running` are restarted |
| `DependentServices` | `@('PrintWorkflowUserSvc','Fax','PrintNotify')` | | `PrintWorkflowUserSvc` is prefix-matched |
| `ForceStopFallback` | `$true` | | Permit terminating a service process stuck in `STOP_PENDING` |

### `Cleanup`

| Key | Default | Notes |
|---|---|---|
| `IncludeExtensions` | `@('.SPL','.SHD')` | **Wildcards are rejected outright.** Never add `'*'` |
| `PurgeLegacyExtensions` | `$false` | Also remove `.SPF` and `~DF*` |
| `Recurse` | `$false` | Must stay false; subdirectories are never enumerated |
| `Backup` | `$false` | Set `$true` in production |
| `BackupPath` | `$null` | `$null` = `%ProgramData%\PrintSpoolerCleanup\Backups` |
| `BackupKeepDays` | `30` | `0` disables pruning |
| `RetryCount` | `3` | Delete attempts per locked file |
| `RetryDelayMs` | `500` | Initial delay; backoff is exponential (500 → 1000 → 2000) |
| `WaitForQuiescenceSeconds` | `0` | Wait until the queue stops changing before stopping the service |
| `VerifyAfterRestart` | `$true` | Re-scan and report files that reappeared |
| `PrinterName` / `ExcludePrinterName` | `@()` | Default filters when not given on the command line |
| `OlderThan` | `$null` | Default age filter, e.g. `'2h'` |

### `Logging`

| Key | Default | Notes |
|---|---|---|
| `Level` | `'INFO'` | `TRACE` `DEBUG` `INFO` `WARN` `ERROR` `CRITICAL` |
| `Path` | `$null` | `$null` = first writable candidate root |
| `RetentionDays` | `30` | `0` disables pruning |
| `MaxSizeMB` | `10` | Rotation threshold |
| `MaxFileCount` | `9` | Generations kept (`.log` + `.1` … `.9`) |
| `JsonLog` | `$false` | JSON Lines for log shippers |
| `EventLog` | `$false` | Windows Event Log, source `PrintSpoolerCleanup` |
| `Timestamps` | `$true` | |
| `SessionStampedFileName` | `$false` | `$true` gives each run its own file; `$false` appends to one rotating file — usually what you want for a scheduled task |

### `Execution`

| Key | Default | Notes |
|---|---|---|
| `ThrottleLimit` | `1` | 1 = sequential. Above 1 needs PowerShell 7 for `-Parallel`; a runspace pool is used on 5.1. Max 64 |
| `Force` | `$false` | |
| `NonInteractive` | `$false` | |
| `ConfirmTimeoutSeconds` | `0` | Apply the default answer after N seconds. `0` = wait forever. Set to e.g. `30` on a technician laptop so an unanswered prompt cannot block a maintenance window |
| `WatchdogMinutes` | `0` | `0` disables the watchdog |
| `SelfElevate` | `$false` | |
| `NoColor`, `Ascii`, `Quiet`, `NoProgress` | `$false` | |
| `MachineReadableJson` | `$false` | stdout becomes exactly one JSON document |
| `ExportFormat` | `@()` | `Json`, `Csv`, `Html`, `Text` |
| `ExportPath` | `$null` | |

### `Update`

| Key | Default | Notes |
|---|---|---|
| `Enabled` | `$true` | |
| `AutoCheck` | `$true` | Quiet, cooldown-limited check at start-up |
| `CooldownHours` | `24` | `0` = check every run |
| `Channel` | `'latest'` | `latest`, `stable`, `prerelease` |
| `IncludePrerelease` | `$false` | |
| `VerifyChecksum` | `$true` | **Do not turn this off.** It means installing unverified bytes |
| `Repository` | `'paulmann/Print-Spooler-Queue-Cleanup'` | Point at an internal mirror if you keep one |
| `Branch` | `'main'` | Used by the raw-file discovery fallback |
| `Token` | `$null` | Prefer `PSQ_GITHUB_TOKEN` over putting a token in a file |

### `Diagnostics`

| Key | Default |
|---|---|
| `IncludeEventLog` | `$true` |
| `EventLogHours` | `72` (1–2160) |
| `EventLogMaxEntries` | `200` |
| `IncludeDriverStore`, `IncludePortMonitor`, `IncludeRegistry` | `$true` |

### `Remote`

| Key | Default | Notes |
|---|---|---|
| `PreFlightProbe` | `$true` | Probe WinRM first: offline hosts fail in ~1 s instead of ~40 s |
| `ProbeTimeoutSeconds` | `10` (1–300) | |
| `SessionReuse` | `$true` | |
| `Authentication` | `'Default'` | |
| `UseSsl` | `$false` | WinRM over 5986 |

---

## Environment variables

All optional. Useful for machine-wide policy via GPO Preferences, or for CI.

| Variable | Effect |
|---|---|
| `PSQ_COMPUTERS` | Comma- or semicolon-separated target list |
| `PSQ_FORCE` | `true`/`1`/`yes` → `-Force` |
| `PSQ_LOG_PATH` | Explicit log file |
| `PSQ_LOG_LEVEL` | `TRACE`…`CRITICAL` |
| `PSQ_LOG_RETENTION_DAYS` | Rotated-log retention |
| `PSQ_LOG_JSON` | JSON Lines log |
| `PSQ_EVENT_LOG` | Also write to the Event Log |
| `PSQ_CONFIG_PATH` | Configuration file to load |
| `PSQ_SERVICE_TIMEOUT` | Service state-change timeout |
| `PSQ_SETTLE_SECONDS` | Post-stop pause |
| `PSQ_RETRY_COUNT` | Delete retries |
| `PSQ_THROTTLE_LIMIT` | Parallel remote workers |
| `PSQ_BACKUP_PATH` | Backup root — **also implies `-Backup`** |
| `PSQ_PRINTER_NAME` | Default printer filter list |
| `PSQ_OLDER_THAN` | Default age filter |
| `PSQ_EXPORT_PATH` | Report destination |
| `PSQ_EXPORT_FORMAT` | `Json,Csv,Html,Text` |
| `PSQ_MACHINE_JSON` | stdout JSON only |
| `PSQ_QUIET` | Suppress console decoration |
| `PSQ_NO_COLOR` | Plain output (also honours `NO_COLOR`) |
| `PSQ_ASCII` | ASCII glyphs |
| `PSQ_NONINTERACTIVE` | Never prompt |
| `PSQ_NO_UPDATE_CHECK` | Disable automatic update checks |
| `PSQ_UPDATE_COOLDOWN_HOURS` | Minimum hours between checks (`0` = always) |
| `PSQ_UPDATE_CHANNEL` | `latest` / `stable` / `prerelease` |
| `PSQ_GITHUB_TOKEN` | Raises the anonymous API limit to 5000/h |

Boolean values accept `1`, `true`, `yes`, `y`, `on`, `enabled` (case-insensitive).

`PSQ_GITHUB_TOKEN` is read from the environment and is never written to the log file, the report, or
a diagnostics bundle. Prefer it over `Update.Token` in a file, and over `-GitHubToken` on a command
line — a token on a command line is visible in process listings and in task definitions.

> **Implementation note:** `PSQ_PRINTER_NAME` is split with `-split '[,;]'`, not with
> `.Split(',', ';')`. The latter binds to `String.Split(char, int)` and silently *truncates* the
> value instead of splitting on both delimiters — a real defect that a unit test caught.

---

## Validation

**Invalid configuration is never silently ignored.** An unknown section, an unknown key, a
non-numeric value in a numeric field, an out-of-range number, an invalid log level, an invalid export
format or an invalid update channel all abort the run with exit code **3** and a message naming the
problem:

```
[ERROR] Configuration: Service.TimeoutSeconds must be between 5 and 3600 (got 99999).
[ERROR] Configuration: Unknown key 'Cleanup.Backkup'.
```

This is deliberate. A silently ignored typo is how a site discovers six months later that its
backups were never on.

### Hard safety rules

Some values are refused regardless of range:

- `Cleanup.IncludeExtensions` may never contain `'*'`, `'.*'` or an empty string — that would turn
  the delete step into "delete everything in the folder"
- `Cleanup.Recurse` must remain `$false`
- Extensions must match `^\.[A-Za-z0-9~]{1,6}$`

---

## Security of the config file

Configuration files travel: they arrive from a network share, a Git repository, a deployment package.
A config that can execute code is a remote-code-execution vector with a friendly name.

The file is parsed with the **PowerShell AST** and evaluated with `SafeGetValue()`. Rejected
outright:

```powershell
# Command invocation in a value
@{ Service = @{ TimeoutSeconds = (Invoke-Expression '1+1') } }

# An executable statement
@{ Service = @{ TimeoutSeconds = 90 } }
Remove-Item C:\ -Recurse -Force

# A pipeline
@{ Service = @{ Name = 'Spooler' } } | Out-Null

# A variable assignment
$x = 1
@{ Service = @{ TimeoutSeconds = $x } }
```

Each produces a specific error naming the offending line, and the run aborts with exit 3. Files
larger than 1 MB are refused, and a malformed file reports a parse error rather than throwing.

---

## Site examples

### Busy print server, conservative

```powershell
@{
    Service = @{ TimeoutSeconds = 180; SettleSeconds = 5 }
    Cleanup = @{ Backup = $true; BackupRequired = $true; BackupKeepDays = 90; OlderThan = '1h' }
    Logging = @{ Level = 'DEBUG'; RetentionDays = 90; MaxSizeMB = 50 }
    Execution = @{ ExportFormat = @('Csv'); ExportPath = '\\filesrv\logs\spooler' }
}
```

### Workstation fleet, unattended

```powershell
@{
    Cleanup  = @{ Backup = $true; BackupKeepDays = 7; OlderThan = '12h' }
    Execution = @{ ThrottleLimit = 16; Quiet = $true; NonInteractive = $true }
    Logging  = @{ Level = 'WARN'; RetentionDays = 14 }
    Update   = @{ AutoCheck = $true; CooldownHours = 168; Channel = 'stable' }
}
```

### Air-gapped environment

```powershell
@{
    Update  = @{ Enabled = $false; AutoCheck = $false }
    Logging = @{ Level = 'INFO'; RetentionDays = 365; SessionStampedFileName = $false }
}
```

### Audited environment

```powershell
@{
    Logging = @{ Level = 'INFO'; EventLog = $true; JsonLog = $true; RetentionDays = 365 }
    Cleanup = @{ Backup = $true; BackupRequired = $true; BackupKeepDays = 365 }
    Execution = @{ ExportFormat = @('Json', 'Html'); ExportPath = '\\audit\spooler' }
}
```

---

## Related pages

- [Logging and Event Log](Logging-and-Event-Log)
- [CLI Reference](CLI-Reference)
- [Safety Model](Safety-Model) — why some values are refused
- [Enterprise Deployment](Enterprise-Deployment)

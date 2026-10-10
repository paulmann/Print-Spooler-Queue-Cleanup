# Cmdlet Reference

The `PrintSpoolerCleanup` module exports 14 commands. They are the same functions the standalone
script defines, so importing the module and dot-sourcing the script give you an identical surface.

```powershell
Import-Module .\PrintSpoolerCleanup
Get-Command -Module PrintSpoolerCleanup
```

| Command | Purpose |
|---|---|
| `Clear-PrintSpoolerQueue` | The main operation |
| `Get-SpoolerCleanupHelp` | The built-in reference |
| `Get-SpoolerCleanupVersion` | Version, build, host and integrity facts |
| `Test-PrintSpoolerHealth` | Non-destructive health probe with a 0–100 score |
| `Get-PrintSpoolerDiagnostic` | Read-only deep inspection |
| `Get-SpoolerCleanupUpdate` | Is a newer version published? |
| `Update-SpoolerCleanup` | Download, verify and install a newer version |
| `Restore-SpoolerBackup` | Forensic restore of a `-Backup` archive |
| `Get-SpoolerBackup` | List backup bundles |
| `Install-SpoolerCleanupTask` | Register the maintenance task |
| `Get-SpoolerCleanupTask` | Inspect the registered task |
| `Start-SpoolerCleanupTask` | Run the task now |
| `Uninstall-SpoolerCleanupTask` | Remove the task |
| `Get-ExitCodeMeaning` | The exit-code contract as data |

`Clear-PrintSpoolerQueue` is documented in full in the [CLI Reference](CLI-Reference). The rest are
below.

---

## Get-SpoolerCleanupHelp

Renders the built-in reference. The parameter table is generated from the **live parameter metadata**
of `Clear-PrintSpoolerQueue`, so it cannot drift away from the code — a self-test asserts that every
parameter is documented and none carries a placeholder description.

```powershell
Get-SpoolerCleanupHelp                            # the whole document
Get-SpoolerCleanupHelp -Section ExitCodes
Get-SpoolerCleanupHelp -Search backup             # every block mentioning backups
Get-SpoolerCleanupHelp -Sections                  # list section names
Get-SpoolerCleanupHelp -Section Parameters -Width 120
```

| Parameter | Type | Notes |
|---|---|---|
| `-Section` | string | `All` *(default)*, `Usage`, `QuickStart`, `Modes`, `Parameters`, `Examples`, `ExitCodes`, `Configuration`, `Environment`, `Logging`, `Reports`, `Remote`, `Scheduling`, `Updates`, `Safety`, `Troubleshooting`, `Diagnostics`, `Version`, `Credits` |
| `-Search` | string | Keep only blocks containing this term |
| `-Width` | int | Wrap width. `0` = console width, capped at 140 |
| `-Sections` | switch | List section names and exit |

**Output:** `System.String`

---

## Get-SpoolerCleanupVersion

```powershell
Get-SpoolerCleanupVersion
Get-SpoolerCleanupVersion -Detailed | ConvertTo-Json -Depth 5
```

| Field | |
|---|---|
| `Tool`, `ProductFamily`, `Version`, `Build` | Identity and build stamp |
| `Author`, `Email`, `Site`, `Repository`, `Wiki`, `License` | Provenance |
| `MinimumPowerShell`, `RecommendedPowerShell` | `5.1` / `7.4` |
| `InstallDirectory` | Where the running copy lives |

With `-Detailed`, additionally: `HostPowerShell`, `PSEdition`, `OperatingSystem`, `IsWindows`,
`IsAdministrator`, `Context`, `Console`, `LogPath`, and `Files[]` — size, SHA-256, last-write time and
Authenticode status of each installed file.

This is what an asset inventory or compliance scan should collect.

---

## Test-PrintSpoolerHealth

Non-destructive. The Spooler is never stopped and no file is touched, so it is safe to poll.

```powershell
Test-PrintSpoolerHealth
Test-PrintSpoolerHealth -ComputerName SRV-PRINT01, SRV-PRINT02 |
    Format-Table ComputerName, Health, Score, Findings -Wrap
Get-Content hosts.txt | Test-PrintSpoolerHealth -ThrottleLimit 16 |
    Where-Object Health -ne 'Healthy' | Export-Csv unhealthy.csv -NoTypeInformation
```

| Parameter | Type | Default | Notes |
|---|---|---|---|
| `-ComputerName` | `string[]` | local | Pipeline input by value and property name |
| `-Credential` | `PSCredential` | — | Alias `-RunAs` |
| `-MaxQueueAgeHours` | `int` (1–720) | `24` | Age above which a host is degraded |
| `-ThrottleLimit` | `int` (1–64) | `1` | Concurrent remote probes |

**Output:** one object per host.

| Field | Notes |
|---|---|
| `Health` | `Healthy` (≥85), `Degraded` (≥60), `Unhealthy` (<60), `Unknown` (unreachable) |
| `Score` | 0–100, reduced by weighted findings |
| `SpoolerState` | Service state |
| `QueueDepth` | `.SPL` + `.SHD` file count |
| `OldestJobHours` | Age of the oldest queued file |
| `SpoolBytes` | Total spool size |
| `SpoolDirectory` | Resolved path |
| `Findings[]` | Plain-English reasons for every deduction |
| `CheckedUtc` | |

Scoring deductions: Spooler not running (−60), Spooler disabled (−15), `spoolsv.exe` not responding
(−35), working set > 512 MB (−15), dependent not running (−10 each), oldest job over the threshold
(−25), queue depth > 50 (−15), spool > 1 GB (−10), printer offline or in error (−8 each), spooler
crash events in the last 72 h (−20), spool directory missing (−20).

`Findings` is the field to surface in a dashboard — a score without a reason is not actionable.

---

## Get-PrintSpoolerDiagnostic

The `-Diagnostic` mode as a cmdlet. Delegates to `Clear-PrintSpoolerQueue -Diagnostic -Force -PassThru`.

```powershell
Get-PrintSpoolerDiagnostic
Get-PrintSpoolerDiagnostic -ComputerName SRV-PRINT01 -ExportFormat Html -ExportPath C:\Temp
Get-PrintSpoolerDiagnostic | Select-Object -ExpandProperty Results | Select-Object -ExpandProperty Jobs
```

| Parameter | Type | Default |
|---|---|---|
| `-ComputerName` | `string[]` | local |
| `-Credential` | `PSCredential` | — |
| `-ExportFormat` | `string[]` | *(console only)* |
| `-ExportPath` | `string` | — |
| `-EventLogHours` | `int` (1–2160) | `72` |

**Output:** the run-summary object. See [Diagnostics Guide](Diagnostics-Guide).

---

## Get-SpoolerCleanupUpdate

```powershell
Get-SpoolerCleanupUpdate
$info = Get-SpoolerCleanupUpdate -Channel stable
if ($info.UpdateAvailable) { Update-SpoolerCleanup -Force }
```

| Parameter | Type | Default |
|---|---|---|
| `-Channel` | `latest` / `stable` / `prerelease` | `latest` |
| `-GitHubToken` | `string` | `PSQ_GITHUB_TOKEN` |
| `-IncludePrerelease` | `switch` | off |

**Output:** `CurrentVersion`, `LatestVersion`, `UpdateAvailable`, `Channel`, `Source`, `TagName`,
`ReleaseName`, `PublishedUtc`, `Prerelease`, `ReleaseNotes`, `Assets[]`, `DownloadUrl`,
`ExpectedSha256`, `ChecksumSource`, `RawUrl`, `RepositoryUrl`, `CheckedUtc`, `Errors[]`.

`Source` tells you which discovery strategy succeeded (`releases/latest`, `tags`, `raw/main`), and
`Errors[]` records why the earlier ones failed — so a rate-limited run explains itself instead of
silently falling through.

See [Updates and Versioning](Updates-and-Versioning).

---

## Update-SpoolerCleanup

```powershell
Update-SpoolerCleanup -Force
Update-SpoolerCleanup -Force -ApplyNow
Update-SpoolerCleanup -WhatIf
```

| Parameter | Type | Default | Notes |
|---|---|---|---|
| `-Channel` | string | `latest` | |
| `-Force` | switch | off | Install even when not newer (re-install / repair) |
| `-ApplyNow` | switch | off | Move staged files into place immediately |
| `-SkipChecksumVerification` | switch | off | Install without a published checksum |
| `-InstallDirectory` | string | script directory | |
| `-GitHubToken` | string | `PSQ_GITHUB_TOKEN` | |
| `-IncludePrerelease` | switch | off | |

Supports `ShouldProcess`, so `-WhatIf` and `-Confirm` work.

**Output:** `Updated`, `Staged`, `FromVersion`, `ToVersion`, `Files[]`, `BackupPath`,
`VerifyResult`, `Method`, `Message`.

`VerifyResult` is `ChecksumMatch`, `ChecksumMismatch`, `NoChecksumPublished`, `SkippedByOperator` or
`NotRequired`.

---

## Restore-SpoolerBackup

```powershell
Restore-SpoolerBackup -Path .\spool-backup_SRV01_20261010-031500.zip
Restore-SpoolerBackup -Path $report.Results[0].BackupPath -ExpectedHash $report.Results[0].BackupHash `
                      -DestinationDirectory C:\Temp\restored
```

| Parameter | Type | Notes |
|---|---|---|
| `-Path` | string | Aliases `-BackupPath`, `-FullName`. Pipeline by property name |
| `-DestinationDirectory` | string | Default: `restored_<computer>_<stamp>` next to the archive |
| `-ExpectedHash` | string | SHA-256 the archive must match |
| `-Overwrite` | switch | Overwrite existing files in the destination |
| `-Force` | switch | Skip the confirmation |

**Output:** `BackupPath`, `Valid`, `ComputerName`, `CreatedUtc`, `ToolVersion`,
`ManifestFileCount`, `FilesRestored`, `Destination`, `Sha256`, `Jobs[]`, `Note`.

> The `Note` field states the limitation every time: **spool files cannot be re-queued by copying
> them back into `PRINTERS`.** A restore recovers the document payload for re-submission from the
> application. See [Backup and Restore](Backup-and-Restore).

Extraction is zip-slip safe: entry names are reduced to their file-name component.

---

## Get-SpoolerBackup

```powershell
Get-SpoolerBackup | Format-Table CreatedUtc, ComputerName, FileCount, SizeBytes, Valid
Get-SpoolerBackup -BackupPath \\filesrv\logs\spooler-backups
```

`-BackupPath` defaults to `%ProgramData%\PrintSpoolerCleanup\Backups`.

**Output:** one object per `spool-backup_*.zip`, newest first: `File`, `FullPath`, `CreatedUtc`,
`ComputerName`, `FileCount`, `SizeBytes`, `Sha256`, `Valid`, `Error`.

`Valid = $false` with a populated `Error` means the archive is not one of ours, or is corrupt.

---

## Install-SpoolerCleanupTask

```powershell
Install-SpoolerCleanupTask -ScriptPath C:\Tools\Clear-PrintSpoolerQueue.ps1 -Schedule Daily -StartTime 03:00 -PassThru
```

Full documentation in [Scheduling and Automation](Scheduling-and-Automation).

| Parameter | Default |
|---|---|
| `-ScriptPath` | *(required)* |
| `-TaskName` | `PrintSpoolerCleanup\Clear Stuck Print Queue` |
| `-Schedule` | `Daily` (`Weekly`, `Hourly`, `AtStartup`, `OnEvent`, `Once`) |
| `-StartTime` | `03:00` |
| `-IntervalHours` | `6` |
| `-ComputerName` / `-ComputerListFile` | local only |
| `-Arguments` | — |
| `-RunAsAccount` | `NT AUTHORITY\SYSTEM` |
| `-EventLogName`, `-EventSource`, `-EventId` | for `-Schedule OnEvent` |
| `-ExecutionTimeLimitMinutes` | `30` (5–1440) |
| `-RunOnlyIfIdle` | off |
| `-NoBackup`, `-NoAutoUpdate` | off — both are **on** by default |
| `-Method` | `Auto` (`Module` or `Xml`) |
| `-PassThru` | off |

**Output:** `$true`/`$false`, or with `-PassThru` an object carrying `TaskName`, `Method`, `Command`,
`Arguments`, `Schedule`, `RunAsAccount`, `ExecutionTimeLimit`, the generated `Xml`, and the
registered `Task`.

`-Method Xml` forces the UTF-16LE XML + `schtasks.exe` path, which is what works on Server Core and
Windows Server 2012 where the `ScheduledTasks` module may be absent.

---

## Get-SpoolerCleanupTask / Start-SpoolerCleanupTask / Uninstall-SpoolerCleanupTask

```powershell
Get-SpoolerCleanupTask | Format-List *
Start-SpoolerCleanupTask -Wait
Uninstall-SpoolerCleanupTask
```

`Get-SpoolerCleanupTask` returns `TaskName`, `Exists`, `State`, `LastRunTime`, `LastTaskResult`,
**`LastTaskResultMeaning`**, `NextRunTime`, `NumberOfMissedRuns`, `Action`.

`LastTaskResultMeaning` translates the numeric result through the documented contract, so `0x1`
reads as `PartialSuccess` rather than being guessed from memory — under v3 that same value meant
*fatal*.

`Start-SpoolerCleanupTask -Wait` blocks until the task finishes and returns
`{ Started, TaskName, ExitCode, Waited }`. `-TimeoutSeconds` (default 600) bounds the wait.

---

## Get-ExitCodeMeaning

The exit-code contract as data:

```powershell
Get-ExitCodeMeaning | Format-Table Code, Name, Description -AutoSize
Get-ExitCodeMeaning -Code 12      # -> ServiceFailure
```

Useful for generating documentation, building a monitoring rule table, or translating a scheduled
task's `LastTaskResult` without hard-coding a lookup.

---

## Importing the module

```powershell
# From a release archive
Expand-Archive .\PrintSpoolerCleanup-4.0.0.zip -DestinationPath C:\Tools\PrintSpoolerCleanup
Import-Module C:\Tools\PrintSpoolerCleanup\PrintSpoolerCleanup\PrintSpoolerCleanup.psd1

# Or from a clone, after building
./tools/Build.ps1
Import-Module ./dist/PrintSpoolerCleanup/PrintSpoolerCleanup.psd1

# Verify
Get-Module PrintSpoolerCleanup | Format-List Name, Version, ExportedFunctions
```

The manifest declares `CompatiblePSEditions = @('Desktop','Core')` and `PowerShellVersion = '5.1'`.

To make it available to every session, copy `PrintSpoolerCleanup\` into
`$env:ProgramFiles\PowerShell\Modules\` (PowerShell 7) or
`$env:SystemRoot\System32\WindowsPowerShell\v1.0\Modules\` (Windows PowerShell), then
`Import-Module PrintSpoolerCleanup` by name.

---

## Dot-sourcing the script instead

```powershell
. .\Clear-PrintSpoolerQueue.ps1
```

Defines the same 14 functions in your session **without running a cleanup** — a destructive operation
as a side effect of dot-sourcing would be unacceptable. You then call
`Clear-PrintSpoolerQueue` explicitly.

The exit code is published via `$global:LASTEXITCODE` rather than `exit`, so dot-sourcing never
closes your shell.

---

## Related pages

- [CLI Reference](CLI-Reference)
- [Installation and Deployment](Installation-and-Deployment)
- [For Developers](For-Developers) — how the module is generated from `src/`

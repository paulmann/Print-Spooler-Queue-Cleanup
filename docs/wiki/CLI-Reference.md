# CLI Reference

Every parameter of `Clear-PrintSpoolerQueue`, grouped by what it is for.

> **The authoritative list is generated from live parameter metadata.** Run
> `.\Clear-PrintSpoolerQueue.ps1 -HelpSection Parameters` to get this table rendered from the code itself,
> which means it cannot drift. This page is the readable version.

```powershell
.\Clear-PrintSpoolerQueue.ps1 [-ComputerName <string[]>] [-Force] [-Backup] [-WhatIf]
                              [-PrinterName <string[]>] [-OlderThan <string>]
                              [-Diagnostic] [-Repair] [-CheckUpdate] [-Update]
                              [-Help] [-HelpSection <name>] [<CommonParameters>]
```

The function supports `ShouldProcess`, so `-WhatIf` and `-Confirm` work as usual. `-WhatIf` runs
the full inventory and plan without stopping the service or deleting anything.

---

## Targeting

| Parameter | Type | Default | Notes |
|---|---|---|---|
| `-ComputerName` | `string[]` | local machine | Pipeline input by value and by property name. Aliases: `-Computer`, `-CN`, `-HostName`. |
| `-ComputerListFile` | `string` | — | Text file, one computer per line. `#` and `;` start comments. Aliases: `-HostsFile`, `-TargetsFile`. |
| `-Credential` | `PSCredential` | caller identity | For remote targets. Never written to logs, reports or bundles. Alias `-RunAs`. |
| `-ThrottleLimit` | `int` (1–64) | `1` | Concurrent remote workers. 1 = sequential, which is the safe default because a spooler restart is disruptive. |
| `-PreFlightProbe` | `bool` | `$true` | Probe WinRM on each remote target first, so offline hosts fail in ~1 s instead of ~40 s. |
| `-UseSsl` | `switch` | off | Use HTTPS (WinRM 5986) instead of HTTP (5985). |

`localhost`, `.`, `127.0.0.1`, `::1`, `(local)`, the local NetBIOS name and the local FQDN are all
recognised as "this machine" and executed locally rather than over WinRM.

Names are de-duplicated across `-ComputerName`, the pipeline and `-ComputerListFile`. Empty and
whitespace entries are skipped.

```powershell
Get-ADComputer -Filter * -SearchBase "OU=Workstations,DC=corp,DC=local" |
    Clear-PrintSpoolerQueue -Force -Backup            # binds on the Name property

Import-Csv .\assets.csv | Clear-PrintSpoolerQueue -Force   # binds on a ComputerName column
```

---

## Selection

By default **every** queued job is a candidate. Any filter narrows that. Filters combine with AND.

| Parameter | Type | Notes |
|---|---|---|
| `-PrinterName` | `string[]` | Wildcards, case-insensitive: `"HP*"`, `"*Kyocera*"`. Aliases `-Printer`, `-Queue`. |
| `-ExcludePrinterName` | `string[]` | Applied after `-PrinterName`. Protects a queue you must not touch. |
| `-JobId` | `long[]` | Exact job ids as shown by `-Diagnostic` or `Get-PrintJob`. Alias `-Id`. |
| `-DocumentName` | `string` | Wildcard match on the document name, e.g. `"*invoice*.pdf"`. |
| `-OwnerName` | `string` | Submitting user, e.g. `"CORP\jdoe"`. Aliases `-Owner`, `-SubmittedBy`. |
| `-OlderThan` | `object` | `30s`, `15m`, `2h`, `7d`, `1w`, `01:30:00`, or a `[TimeSpan]`. Alias `-MinAge`. |
| `-NewerThan` | `object` | Same formats. Alias `-MaxAge`. Must be smaller than `-OlderThan`. |
| `-LargerThanBytes` | `object` | Byte count or `"10MB"`, `"1.5GB"`. Alias `-MinSize`. |
| `-SmallerThanBytes` | `object` | Byte count or human size. Alias `-MaxSize`. |
| `-SpoolDirectory` | `string` | Override the resolved spool directory. Still subject to every guard. |

**`-OlderThan` is the most valuable filter in production.** It clears the wedged jobs and leaves
the one that is legitimately printing right now.

```powershell
.\Clear-PrintSpoolerQueue.ps1 -OlderThan 2h -Force -Backup
.\Clear-PrintSpoolerQueue.ps1 -PrinterName "HP*","Canon*" -ExcludePrinterName "HP Legal" -Force
.\Clear-PrintSpoolerQueue.ps1 -LargerThanBytes 50MB -Force     # the 4 GB CAD job
```

> **Metadata availability.** Filters need job metadata. While the Spooler answers WMI, metadata is
> authoritative. When it is hung, the inventory comes from `.SHD` headers, where printer and
> document names are best-effort. Jobs with no readable metadata are **not** selected by
> name-based filters, and the run reports this as a warning rather than guessing. Use `-JobId`
> (derived from the file name, always available) or no filter at all when the whole queue must go.

---

## Modes

Exactly one mode is active per run. They are mutually exclusive by design, so a scheduled task can
never accidentally become a destructive one.

| Mode | Effect | Exit code on success |
|---|---|---|
| *(default)* | Stop → delete selected files → restart → verify | 0 |
| `-WhatIf` | Full inventory and plan, zero changes, service not stopped | 0 |
| `-Diagnostic` | Read-only deep inspection | 11 |
| `-Repair` | Cleanup **plus** root-cause repair steps | 0 |
| `-Restore <zip>` | Forensic restore of a `-Backup` archive | 0 / 13 |
| `-CheckUpdate` | Query GitHub and exit | 0 / 6 / 7 |
| `-Update` | Download, verify, stage and optionally apply | 0 / 8 / 9 |
| `-SelfTest` | Internal consistency checks | 0 / 2 |
| `-Help` | Render the complete built-in reference | 0 |
| `-HelpSection <name>` | Render one section of the reference | 0 |
| `-ShowVersion` | Version, build, host, OS, integrity facts | 0 |
| `-SaveConfig` | Write the effective configuration to `-ConfigPath` | 0 / 3 |

Supporting switches: `-AutoUpdate` (quiet cooldown-limited check before a cleanup),
`-Watchdog`, `-SelfTestPrint`, `-ListHelpSections`.

---

## Safety

| Parameter | Type | Default | Notes |
|---|---|---|---|
| `-Force` | `switch` | off | Skip the confirmation prompt. Required for unattended runs. |
| `-Backup` | `switch` | off | Archive every file before deleting it. |
| `-BackupPath` | `string` | `%ProgramData%\PrintSpoolerCleanup\Backups` | Requires `-Backup`. |
| `-BackupRequired` | `switch` | off | Abort if the backup could not be created. |
| `-ArchivePath` | `string` | — | Move created archives here afterwards (e.g. a log share). |
| `-RetryCount` | `int` (0–20) | `3` | Delete attempts per locked file. |
| `-RetryDelayMs` | `int` (0–60000) | `500` | Initial delay; backoff is exponential (500 → 1000 → 2000). |
| `-VerifySignature` | `switch` | off | Verify Authenticode before doing anything else. Exit 14 on failure. |
| `-ExpectedSignerThumbprint` | `string` | — | Require the signature to come from this certificate. |
| `-NonInteractive` | `switch` | auto | Never prompt; abort with exit 4 instead of waiting. Alias `-NoPrompt`. |
| `-WatchdogMinutes` | `int` (0–1440) | `30` | With `-Watchdog`: terminate and restart the Spooler past this. |

`-NonInteractive` is applied automatically when the host cannot accept a prompt
(`[Console]::IsInputRedirected`, Task Scheduler, CI). This is the fix for a scheduled task that
otherwise blocks forever on a Y/N question nobody can answer.

---

## Repair options

All require `-Repair`.

| Parameter | Effect |
|---|---|
| `-PurgeFormsCache` | Clear the spooler Forms cache — a corrupt custom form can wedge the service. |
| `-PurgeDriverStore` | Remove printer driver packages that **no** printer references. Destructive, so it also requires `-Force`. |
| `-ResetSpoolDirectory` | Remove the `DefaultSpoolDirectory` registry override so the OS default is used again. |
| `-StopDependentServices` | `[bool]`, default `$true`. Stop `Fax`/`PrintNotify`/`PrintWorkflowUserSvc_*` first. |
| `-RestartDependentServices` | `[bool]`, default `$true`. Restart only those that were running. |

---

## Output

| Parameter | Type | Default | Notes |
|---|---|---|---|
| `-Quiet` | `switch` | off | Suppress decorative console output. Logging and reports unaffected. Alias `-Silent`. |
| `-NoColor` | `switch` | auto | Monochrome. Also triggered by `NO_COLOR` or a redirected stream. |
| `-Ascii` | `switch` | auto | ASCII glyphs. Auto-enabled when the console code page cannot render Unicode. Alias `-NoUnicode`. |
| `-NoProgress` | `switch` | auto | Suppress progress bars. |
| `-MachineReadableJson` | `switch` | off | Exactly one JSON document on stdout, nothing else. Alias `-Json`. |
| `-ExportFormat` | `string[]` | — | `Json`, `Csv`, `Html`, `Text`. Alias `-ReportFormat`. |
| `-ExportPath` | `string` | — | Directory (auto-named files) or full file path. Alias `-ReportPath`. |
| `-PassThru` | `switch` | off | Return the run summary object to the pipeline. |
| `-DiagnosticsBundlePath` | `string` | — | Also write a ZIP suitable for a support issue. |

---

## Logging

| Parameter | Type | Default | Notes |
|---|---|---|---|
| `-LogPath` | `string` | first writable root | Parent directory is created. Alias `-Log`. |
| `-LogLevel` | `string` | `INFO` | `TRACE`, `DEBUG`, `INFO`, `WARN`, `ERROR`, `CRITICAL`. |
| `-JsonLog` | `switch` | off | JSON Lines instead of pipe-delimited text. |
| `-EventLog` | `switch` | off | Also write to the Windows Event Log (source `PrintSpoolerCleanup`). |
| `-LogRetentionDays` | `int` (0–3650) | `30` | Delete rotated logs older than this. 0 disables pruning. |
| `-NoLogRotation` | `switch` | off | Disable size-based rotation for this run. |

See [Logging and Event Log](Logging-and-Event-Log).

---

## Configuration

| Parameter | Type | Default | Notes |
|---|---|---|---|
| `-ConfigPath` | `string` | auto-discovered | Load a `.psd1` configuration file. Alias `-Config`. |
| `-SaveConfig` | `switch` | off | Write the effective configuration to `-ConfigPath` and exit. |
| `-NoConfigFile` | `switch` | off | Ignore any auto-discovered configuration file. |
| `-ServiceTimeoutSeconds` | `int` (5–3600) | `60` | Seconds to wait for a service state change. Alias `-TimeoutSeconds`. |
| `-SettleSeconds` | `int` (0–300) | `2` | Pause after the Spooler stops so handles are released. |
| `-EventLogHours` | `int` (0–2160) | `72` | Event-log history collected by `-Diagnostic`. |

See [Configuration](Configuration).

---

## Updates

| Parameter | Type | Default | Notes |
|---|---|---|---|
| `-UpdateChannel` | `string` | `latest` | `latest`, `stable` (no prereleases), `prerelease`. |
| `-IncludePrerelease` | `switch` | off | Treat prereleases as valid targets. |
| `-SkipChecksumVerification` | `switch` | off | Install even when no `checksums.txt` is published. |
| `-GitHubToken` | `string` | `PSQ_GITHUB_TOKEN` | Raises the anonymous 60 req/h limit to 5000. |
| `-NoUpdateCheck` | `switch` | off | Disable every automatic check for this run. |
| `-ApplyUpdateNow` | `switch` | off | Move staged files into place immediately. |

Prefer `PSQ_GITHUB_TOKEN` over `-GitHubToken` so the secret never appears in a command line or a
process listing. See [Updates and Versioning](Updates-and-Versioning).

---

## Environment

| Parameter | Notes |
|---|---|
| `-SelfElevate` | Re-launch elevated through UAC when not running as administrator. Alias `-Elevate`. |
| `-NoElevate` | Never attempt elevation; fail with exit 5. Mutually exclusive with `-SelfElevate`. |
| `-NoProfile` | Accepted and ignored, for compatibility with launchers that always pass it. |

---

## Help

| Parameter | Notes |
|---|---|
| `-Help` | Render the complete built-in reference and exit. |
| `-HelpSection <name>` | Render one section and exit. `-ListHelpSections` prints the names. |
| `-ListHelpSections` | Print the section names and exit. |
| `-ShowVersion` | Version, build, host, OS, elevation and file integrity. |

---

## Parameter conflicts

These combinations are rejected with exit code **3** and an explanation, rather than being silently
resolved:

- `-Restore` with `-Repair`, or with any selection filter
- `-Diagnostic` with `-Repair`, `-PrinterName` or `-JobId`
- `-JobId` with `-PrinterName` (a job id is already exact)
- `-Update` with `-ComputerName` (the tool updates itself, not a remote host)
- `-SelfElevate` with `-NoElevate`
- `-AutoUpdate` with `-CheckUpdate`
- `-SaveConfig` without `-ConfigPath`
- `-BackupPath` without `-Backup`
- `-MachineReadableJson` with `-ExportFormat`
- `-OlderThan` ≤ `-NewerThan` (nothing could ever match)
- `-PurgeDriverStore` without `-Force`

---

## Common parameters

Because the function is `Advanced`, the standard set is available:
`-Verbose`, `-Debug`, `-ErrorAction`, `-WarningAction`, `-InformationAction`, `-OutVariable`,
`-OutBuffer`, `-PipelineVariable`, `-WhatIf`, `-Confirm`.

`-Verbose` is genuinely useful during `-Repair`: it echoes every log record to the verbose stream
in addition to the file.

---

## See also

- [Cmdlet Reference](Cmdlet-Reference) — the 14 module cmdlets
- [Batch Launcher Reference](Batch-Launcher-Reference) — the `.bat` equivalents
- [Exit Codes](Exit-Codes)
- [Safety Model](Safety-Model)

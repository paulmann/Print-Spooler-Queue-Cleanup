# Exit Codes

The exit code is the only thing most monitoring systems will ever read, so it is a **contract**:
stable, documented, identical for the `.ps1` and the `.bat`, and enforced by the test suite.

Retrieve it with `$LASTEXITCODE` in PowerShell, `%ERRORLEVEL%` in cmd, or the process exit status
from any scheduler or RMM agent.

---

## The contract

| Code | Name | Meaning | Severity |
|---:|---|---|---|
| **0** | `Success` | Every target processed and the Spooler verified running. | Healthy |
| **1** | `PartialSuccess` | Some targets succeeded; others failed or produced warnings. | Investigate |
| **2** | `Failed` | No target could be processed. | **Alert** |
| **3** | `BadArguments` | Invalid parameter combination, unreadable computer list, or a configuration file that failed validation. | Fix the caller |
| **4** | `Cancelled` | The operator declined the prompt, or the host is non-interactive and `-Force` was not supplied. | Fix the caller |
| **5** | `NotAdministrator` | Elevation required and not available. | Fix the caller |
| **6** | `UpdateCheckFailed` | `-CheckUpdate` could not reach GitHub (offline, proxy, rate limit). | Informational |
| **7** | `UpdateAvailable` | `-CheckUpdate` found a newer version. **Nothing was changed.** | Plan an update |
| **8** | `UpdateApplied` | `-Update` / `-AutoUpdate` installed a newer version. | Informational |
| **9** | `UpdateFailed` | Download, checksum verification or staging failed. | Investigate |
| **10** | `UnsupportedPlatform` | Not Windows, or the Print Spooler feature is not installed. | Fix the target list |
| **11** | `DiagnosticsOnly` | `-Diagnostic` completed. Informational — **not an error**. | Healthy |
| **12** | `ServiceFailure` | The Spooler could not be stopped or restarted. | **Alert** |
| **13** | `RestoreFailed` | `Restore-SpoolerBackup` could not complete. | Investigate |
| **14** | `SignatureInvalid` | `-VerifySignature` found a broken, missing or unexpected signature. | **Alert** |
| **15** | `Timeout` | The watchdog expired and the run was terminated. | **Alert** |
| **130** | `Aborted` | Interrupted with Ctrl+C or by host shutdown. | Investigate |

---

## Building monitoring rules

**Page a human on:** `2`, `5`, `12`, `14`, `15`

- `2` — nothing was cleaned; the queue is still stuck and users are still blocked.
- `12` — the Spooler could not be stopped or restarted. **Printing may be down.** This is the most
  urgent code in the set.
- `5` — the task or shortcut is not elevated, so maintenance has been silently not happening.
- `14` — the file on disk does not match its signature. Treat as a possible tampering incident.
- `15` — the run wedged badly enough that the watchdog killed it.

**Open a ticket, do not page:** `1`, `9`, `13`

- `1` — some hosts worked, some did not. Usually a handful of offline machines or one bad driver.
  The per-computer table in the report says which.
- `9` — the update did not apply. The installed version still works; you are simply not current.
- `13` — a restore failed. Someone is waiting on recovered documents.

**Informational:** `0`, `6`, `7`, `8`, `11`

- `7` is worth a *dashboard tile*, not an alert: "N of your print servers are running a stale
  cleanup tool" is useful context during an incident, not a reason to wake anyone.
- `6` in an air-gapped environment is expected and permanent — set `PSQ_NO_UPDATE_CHECK=true` and
  it stops appearing.

**Caller bugs:** `3`, `4`, `10`

These almost always mean the invocation is wrong, not that the print server is broken. A scheduled
task returning `4` was built by hand and is missing `-NonInteractive`; a task returning `3` has a
typo in its arguments.

---

## Mapping from result status

Each per-computer result carries a `Status` string; the run-level code is derived from them:

| Result statuses present | `OverallStatus` | Exit code |
|---|---|---:|
| all `Success` | `Success` | 0 |
| all `NothingToDo` | `NothingToDo` | 0 |
| all `DryRun` | `DryRun` | 0 |
| all `Diagnostic` | `Diagnostic` | 11 |
| any `SuccessWithWarnings`, none failed | `SuccessWithWarnings` | 1 |
| some succeeded, some `Failed`/`ServiceFailure` | `PartialSuccess` | 1 |
| all failed | `Failed` | 2 |
| all `Cancelled`/`Skipped` | `Cancelled` | 4 |
| any `ServiceFailure` (overrides the above, up to 12) | — | 12 |

`NothingToDo` maps to **0**, not to a warning: "the queue was already empty" is a successful
outcome for a scheduled maintenance run, and treating it otherwise produces a permanent sea of
amber in your dashboard.

---

## Version note (breaking change from v3)

v3.1.0 used a five-code contract:

| v3 | v4 equivalent |
|---|---|
| `0` success | `0` Success |
| `1` fatal error | `2` Failed |
| `2` completed with warnings | `1` PartialSuccess |
| `3` bad arguments | `3` BadArguments |
| `4` cancelled by user | `4` Cancelled |

**Codes 1 and 2 swapped meaning.** If you have monitoring rules written against v3, update them:
a v3 rule alerting on `1` (fatal) would, under v4, fire on a partially successful run and stay
silent on a total failure.

The `Get-ExitCodeMeaning` cmdlet returns the whole contract as data, which is the safest way to
generate a mapping table:

```powershell
Get-ExitCodeMeaning | Format-Table Code, Name, Description -AutoSize
Get-ExitCodeMeaning -Code 12      # -> ServiceFailure
```

---

## Scheduled tasks

`Get-SpoolerCleanupTask` reports `LastTaskResult`, and adds a `LastTaskResultMeaning` property that
translates it through this same contract:

```powershell
Get-SpoolerCleanupTask | Format-List TaskName, State, LastRunTime, LastTaskResult, LastTaskResultMeaning, NextRunTime
```

A `LastTaskResult` of `0x1` therefore means *PartialSuccess*, not "fatal" — which is precisely the
ambiguity that the `LastTaskResultMeaning` property exists to remove.

---

## See also

- [CLI Reference](CLI-Reference)
- [Reports and Machine-Readable Output](Reports-and-Machine-Readable-Output) — the JSON contract
  carries `exitCode` inside the document as well as in the process status
- [Troubleshooting](Troubleshooting)

# Exit Code Contract

**Status:** normative. Enforced by `tests/Unit.Utilities.ps1` and asserted by the batch-lint stage of
`tests/Run-Tests.ps1`.

The exit code is the only thing most monitoring systems will ever read, so it is a contract: stable,
documented, identical for `Clear-PrintSpoolerQueue.ps1` and `Clear-PrintSpoolerQueue.bat`, and
test-enforced. Changing a code's meaning is a breaking change requiring a major version bump.

---

## Retrieval

| Context | How |
|---|---|
| PowerShell, after `-File` or direct invocation | `$LASTEXITCODE` |
| PowerShell, after dot-sourcing | `$global:LASTEXITCODE` (`exit` is deliberately not used, so dot-sourcing never closes the shell) |
| `cmd.exe` | `%ERRORLEVEL%` |
| Task Scheduler | `LastTaskResult` — `Get-SpoolerCleanupTask` adds a `LastTaskResultMeaning` property |
| RMM / CI agent | process exit status |
| In-process | `(-PassThru).ExitCode`, or `summary.exitCode` in the JSON document |

---

## The contract

| Code | Name | Meaning | Severity |
|---:|---|---|---|
| **0** | `Success` | Every target processed; Spooler verified running. Also returned for `NothingToDo` and `DryRun` | Healthy |
| **1** | `PartialSuccess` | ≥1 target succeeded and ≥1 failed or warned, **or** a run succeeded with warnings | Investigate |
| **2** | `Failed` | No target could be processed | **Alert** |
| **3** | `BadArguments` | Invalid parameter combination, unreadable computer list, or a configuration file that failed validation | Caller bug |
| **4** | `Cancelled` | The operator declined the prompt, or the host is non-interactive and `-Force` was not supplied | Caller bug |
| **5** | `NotAdministrator` | Elevation required and not available | Caller bug |
| **6** | `UpdateCheckFailed` | `-CheckUpdate` could not reach GitHub (offline, proxy, rate limit) | Informational |
| **7** | `UpdateAvailable` | `-CheckUpdate` found a newer version. **Nothing was changed** | Dashboard |
| **8** | `UpdateApplied` | `-Update` / `-AutoUpdate` installed a newer version | Informational |
| **9** | `UpdateFailed` | Download, checksum verification or staging failed | Investigate |
| **10** | `UnsupportedPlatform` | Not Windows, or the Print Spooler feature is not installed | Targeting bug |
| **11** | `DiagnosticsOnly` | `-Diagnostic` completed. Informational — **not an error** | Healthy |
| **12** | `ServiceFailure` | The Spooler could not be stopped or restarted | **Alert** |
| **13** | `RestoreFailed` | `Restore-SpoolerBackup` could not complete | Investigate |
| **14** | `SignatureInvalid` | `-VerifySignature` found a broken, missing or unexpected signature | **Alert** |
| **15** | `Timeout` | The watchdog expired and the run was terminated | **Alert** |
| **130** | `Aborted` | Interrupted with Ctrl+C or by host shutdown | Investigate |

Codes are also available as data:

```powershell
Get-ExitCodeMeaning | Format-Table Code, Name, Description -AutoSize
Get-ExitCodeMeaning -Code 12      # -> ServiceFailure
```

---

## Derivation

Each per-computer result carries a `Status`. `New-RunSummary` aggregates them into an
`OverallStatus`, and `Get-ExitCodeForSummary` maps that to a code:

| Result statuses present | `OverallStatus` | Exit |
|---|---|---:|
| all `Success` | `Success` | 0 |
| all `NothingToDo` | `NothingToDo` | 0 |
| all `DryRun` | `DryRun` | 0 |
| all `Diagnostic` | `Diagnostic` | 11 |
| any `SuccessWithWarnings`, none failed | `SuccessWithWarnings` | 1 |
| some succeeded, some `Failed`/`ServiceFailure` | `PartialSuccess` | 1 |
| all failed | `Failed` | 2 |
| all `Cancelled`/`Skipped` | `Cancelled` | 4 |
| zero results | `Failed` | 2 |

Two overrides apply afterwards:

1. Any `ServiceFailure` raises the code to **12** — a host left without a Spooler outranks a partial
   success.
2. `-Diagnostic` with an otherwise-successful outcome becomes **11**.

Codes 3, 4, 5, 6, 7, 8, 9, 10, 13, 14 and 15 are set directly by the corresponding code path and do
not derive from per-target results.

`NothingToDo` maps to **0** deliberately: for a scheduled maintenance run, "the queue was already
clear" is success. Treating it otherwise produces a permanent sea of amber in a dashboard, which
trains people to ignore it.

---

## Monitoring policy

**Page a human:** `2`, `5`, `12`, `14`, `15`

| Code | Why it is worth waking someone |
|---|---|
| 2 | Nothing was cleaned. The queue is still stuck and users are still blocked |
| 12 | The Spooler could not be stopped or restarted. **Printing may be down** |
| 5 | The task is not elevated, so maintenance has been silently not happening |
| 14 | The file on disk does not match its signature — treat as a possible tampering incident |
| 15 | The run wedged badly enough that the watchdog killed it |

**Ticket, do not page:** `1`, `9`, `13`

**Dashboard only:** `6`, `7`. "N of your print servers run a stale cleanup tool" is useful context
during an incident, not a reason to wake anyone. In an air-gapped site, set `PSQ_NO_UPDATE_CHECK=true`
and `6` stops appearing.

**Healthy:** `0`, `8`, `11`.

**Caller bugs:** `3`, `4`, `10`. These mean the *invocation* is wrong, not that the print server is
broken. A scheduled task returning `4` was hand-made and is missing `-NonInteractive`; one returning
`3` has a typo in its arguments.

---

## Scheduled tasks

```powershell
Get-SpoolerCleanupTask | Format-List TaskName, State, LastRunTime, LastTaskResult, LastTaskResultMeaning, NextRunTime
```

| `LastTaskResult` | Meaning | Action |
|---|---|---|
| `0x0` | Success | It worked; check the log for what it did |
| `0x1` | PartialSuccess | Some targets failed — read the report |
| `0x2` | Failed | Read the log; usually a guard or a service failure |
| `0x4` | Cancelled | **It waited for a confirmation nobody could give.** Add `-Force -NonInteractive` |
| `0x5` | NotAdministrator | Run as `SYSTEM` or an elevated account |

`0x4` is the classic hand-made-task failure, which is why `Install-SpoolerCleanupTask` always adds
`-Force -NonInteractive -Quiet -NoProgress`.

---

## Breaking change from v3

v3.1.0 used a five-code contract. **Codes 1 and 2 swapped meaning.**

| v3 | v4 | |
|---|---|---|
| `0` success | `0` Success | unchanged |
| `1` fatal error | `2` Failed | **moved** |
| `2` completed with warnings | `1` PartialSuccess | **moved** |
| `3` bad arguments | `3` BadArguments | unchanged |
| `4` cancelled by user | `4` Cancelled | unchanged |

A v3 monitoring rule that alerted on `1` would, under v4, fire on a partially successful run and stay
silent on a total failure. **Update any rule keyed on these codes.**

---

## Implementation notes

- `$global:LASTEXITCODE = N` alone does **not** change the process exit code for
  `powershell -File`. Only `exit N` does — and `exit N` from a dot-sourced script closes the
  operator's shell. `Publish-ScriptExitCode` therefore branches on the invocation context, which must
  be captured **in script scope** because `$MyInvocation` inside a function describes the function,
  not the script. Getting this wrong is the v3 defect where `-File` reported success after a failure.
- The table lives in `$Script:ExitCode` in `src/Shared/00-Constants.ps1` and is the single source for
  the script, the launcher help text, the built-in `-HelpSection ExitCodes`, the README, the wiki and
  `Get-ExitCodeMeaning`.
- The batch-lint stage asserts that codes 2, 3, 5 and 10 are actually reachable in the launcher, so a
  documented code cannot become dead documentation.

---

## See also

- [Wiki: Exit Codes](../wiki/Exit-Codes.md) — the operator-facing version with remediation steps
- [Report Schema](./Report-Schema.md) — `exitCode` also appears inside the JSON document
- [Wiki: Troubleshooting](../wiki/Troubleshooting.md)

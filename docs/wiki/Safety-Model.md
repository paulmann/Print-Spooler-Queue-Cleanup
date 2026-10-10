# Safety Model

This tool deletes files as an administrator on production print servers. Everything on this page
exists because the alternative is a self-inflicted outage.

If you read only one wiki page, read this one.

---

## The one destructive operation

There is exactly **one** place in the product where a file is deleted: inside the execution payload
that runs on the target machine. The same payload is invoked directly for the local machine and
through `Invoke-Command` for remote ones, so local and remote behaviour cannot drift.

The batch launcher does not implement its own cleanup. v3 did, and that meant two destructive code
paths for one job — two chances to be wrong, and a batch copy that could never offer remote,
selective or backed-up cleanup. v4 delegates, so the question "what can this tool delete?" has one
answer that can be audited in one place.

---

## The four guards

Before a single file is removed, the resolved spool directory must pass four independent checks.
**If any one fails, nothing is deleted** and the run exits `2` (Failed) with a message naming the
guard that fired.

### 1. Extension allow-list

Only `.SPL` and `.SHD` are deleted. Optionally `.SPF` and `~DF*` with
`Cleanup.PurgeLegacyExtensions = $true`.

The list can never contain a wildcard: a configuration with `IncludeExtensions = @('*')` or `@('.*')`
is rejected at validation time with exit code 3. Deletion enumerates files only — never directories.

### 2. Path validation

The resolved directory must:

- be rooted at a Windows drive letter (`^[A-Za-z]:[\\/]`)
- contain no characters illegal in a Windows path (`< > : " | ? *`, other than the drive colon)
- contain no control characters
- **not** be a UNC path — remote file systems must be cleaned on their owner
- end with the leaf `PRINTERS` (case-insensitive)
- be at least 3 path segments deep
- live under a folder named `spool`

That last rule is the one that does the most work. A corrupt `DefaultSpoolDirectory` pointing at
`C:\Users\PRINTERS` has the right leaf *and* the right depth, and an equality-only forbidden-root
check would accept it. Requiring a real `spool` folder in the path defeats that whole class of
accident.

> **Why the drive-letter check comes before `GetFullPath()`:** `GetFullPath` is *platform-relative*.
> On a non-Windows host, `C:\PRINTERS` resolves to `<cwd>/C:\PRINTERS` — eleven segments deep, which
> would silently satisfy the depth guard. Validating the shape first makes the guard correct on
> every platform rather than only on the one it was written on. This was found by running the test
> suite on Linux, which is exactly why the CI matrix includes it.

### 3. Protected roots

The resolved path must not equal a protected system root: `C:\`, `C:\Windows`,
`C:\Windows\System32`, `C:\Windows\SysWOW64`, `C:\Users`, `C:\Program Files`,
`C:\Program Files (x86)`, `C:\ProgramData`, `C:\Recovery`, `C:\PerfLogs`, `C:\Boot`, plus
`%SystemRoot%`, `%windir%`, `%ProgramFiles%`, `%ProgramFiles(x86)%`, `%ProgramData%`,
`%USERPROFILE%` and `%SystemDrive%\` resolved at runtime.

Machine-specific values are resolved at run time rather than baked in, so an unset environment
variable can never inject `$null` into the comparison set.

### 4. Volume guards

More than **100 000** files or more than **100 GB** in the directory means "this is not a spool
directory". The run aborts with an error telling you to investigate with `-Diagnostic`.

A real print server queue is dozens of files and hundreds of megabytes at worst. Six-figure counts
only appear when something has resolved to the wrong tree, and at that point deleting is the last
thing anyone wants.

### No recursion

The delete enumeration is `-File` without `-Recurse`. Subdirectories of `PRINTERS` are never
walked. A test asserts that no `Remove-Item -Recurse` anywhere in the payload targets the spool
directory — the only recursive deletes permitted are the tool's own GUID-named backup staging
folders.

---

## The registry override is validated, not trusted

`HKLM\SYSTEM\CurrentControlSet\Control\Print\Printers\DefaultSpoolDirectory` can legitimately
relocate the spool directory, so it is honoured. But it can also be set to anything. A value that
is a UNC path, shorter than 4 characters, or contains control characters is **ignored with a
warning** and the OS default is used instead — such a value is virtually always a misconfiguration,
and acting on it would be the worst possible response to finding one.

The resolved path and its source (`Registry`, `OSDefault`, or `Parameter`) appear in the execution
plan before anything happens, and in every report afterwards.

---

## Service handling

| Rule | Why |
|---|---|
| Dependents (`Fax`, `PrintNotify`, `PrintWorkflowUserSvc_*`) are stopped **before** the Spooler | They hold handles to spool files; this is the single most common reason a stop hangs |
| Only services that were actually `Running` are restarted | A disabled Fax service must not be silently enabled by a cleanup |
| Start types are never modified | The tool is not a configuration manager |
| A service stuck in `STOP_PENDING` can be escalated to a process kill | `ForceStopFallback`. Safe for the Spooler — the SCM treats an exited service process as stopped |
| If the Spooler cannot be stopped, **nothing is deleted** | The run returns `ServiceFailure` (exit 12) and the service is left as it was found |
| Any unexpected error triggers a last-resort restart | A failed cleanup must never become "printing is disabled" |
| Every state transition is recorded | `serviceTrace[]` in the report: service, action, from, to, success, duration |

The `-SettleSeconds` pause (default 2 s) after the stop is not decorative. It lets `spoolsv.exe`
release its handles, which is the difference between zero locked files and a dozen.

---

## Backups

`-Backup` archives the exact set of files it is about to delete, plus a `backup-manifest.json`
recording the tool version, timestamp, computer, spool directory, per-file names and sizes, the job
metadata, and a SHA-256 of the archive.

`-BackupRequired` turns a backup failure into a hard abort — use it wherever a print queue has legal
or financial significance.

### The honest limitation

**Restoring spool files into `PRINTERS` does not re-queue the job.** The Spooler builds its job
table in memory and only reads `.SHD`/`.SPL` files for jobs it created during the current boot.
A restore is a *forensic* recovery of the document payload so it can be re-submitted from the
application — not a transactional rollback.

`Restore-SpoolerBackup` says this in its output, the manifest says it in `RestoreNote`, and the HTML
report says it under the backup table. Anything claiming otherwise would be a lie in your runbook,
and a lie in a runbook gets someone paged at 3 a.m.

See [Backup and Restore](Backup-and-Restore).

---

## Confirmation

Interactive by default. `-Force` skips the prompt.

On a **non-interactive host** — Task Scheduler, CI, redirected stdin — the tool refuses to prompt
and aborts with exit code **4** rather than blocking forever. v3's batch file hung indefinitely in
exactly this situation, which is why the README's own `schtasks` example never completed.

The prompt states which promise is being made:

```
Stop the Print Spooler, delete ALL spool files, then restart the service [Y/N]
Stop the Print Spooler, delete the SELECTED spool files, then restart the service [Y/N]
```

"ALL" versus "SELECTED" is the distinction an operator needs to see before pressing Y.
`Execution.ConfirmTimeoutSeconds` can apply a default answer after a timeout, so an unanswered
prompt cannot block a maintenance window.

---

## Configuration is data, never code

`PrintSpoolerCleanup.config.psd1` is parsed with the PowerShell AST and evaluated with
`SafeGetValue()`. Command invocations, multi-element pipelines and variable assignments are
**rejected outright**, and files larger than 1 MB are refused.

That matters because config files travel: they come from a network share, a Git repository, a
deployment package. A config that can execute code is a remote-code-execution vector with a
friendly name.

Invalid values are never silently ignored. An unknown section, an unknown key, a non-numeric value
in a numeric field, an out-of-range number, an invalid log level or an invalid export format all
abort the run with exit code **3** and a message naming the problem. A silently ignored setting is
how a site discovers six months later that its backups were never on.

---

## Selection filters fail closed, not open

A selection filter narrows a destructive run: `-PrinterName 'HP*'` means "only HP
queues", `-ExcludePrinterName 'PDF'` means "never the PDF writer". The dangerous
failure mode for such a filter is not that it matches too much by accident — it is
that it is **silently discarded**, leaving the run unfiltered.

That is exactly what an empty value does in PowerShell, because `''` is *falsy*:

```powershell
# the payload's filter gate
if ($Options['PrinterName'] -and @($Options['PrinterName']).Count -gt 0) { ... }
```

With `-PrinterName ''` the condition is false, the filter is skipped, and the run
cleans **every** printer and every job. The mirror case is worse: a blank
`-ExcludePrinterName` withdraws a protection the operator explicitly asked for, so
a queue that was meant to be left alone gets emptied.

This was reachable from the launcher — typing `/printer` and forgetting the value
forwarded `-PrinterName ""`.

So blank values are **rejected**, not interpreted:

| Input | Result |
|---|---|
| `-PrinterName 'HP*'` | filter applied |
| `-PrinterName ''` | **exit 3**, run refused |
| `-PrinterName '   '` | **exit 3**, run refused |
| `-PrinterName @('','')` | **exit 3**, run refused |
| `-PrinterName` omitted | no filter — the operator never asked for one |

The distinction that matters is the last row: *omitting* a filter is a legitimate
choice and means "all printers", while passing a *blank* one is a mistake and is
refused. Conflating the two is what turns a typo into a fleet-wide deletion.

The rejection message states the consequence rather than just the rule, because an
operator who does not understand why a blank is dangerous will "fix" it by deleting
the switch and getting the same unfiltered run:

```
-PrinterName was given a blank value (''). A blank selection filter is falsy in
PowerShell, so it would be IGNORED rather than matched - silently widening the run
to EVERY printer and every job. Specify a real value, or omit -PrinterName entirely.
```

Thirteen parameters are covered — the six selection filters (`PrinterName`,
`ExcludePrinterName`, `OwnerName`, `DocumentName`, `OlderThan`, `NewerThan`) and
seven path parameters (`BackupPath`, `ArchivePath`, `ExportPath`, `ConfigPath`,
`RestoreDestination`, `ComputerListFile`, `DiagnosticsBundlePath`), where a blank
would silently fall back to a default location the operator did not choose. The
list lives in one place, `$Script:BlankValueIsUnsafeParameters` in
`src/Shared/00-Constants.ps1`, so adding a future selection parameter forces a
decision about whether it belongs there.

The `.bat` launcher refuses independently: a value-taking switch given no value
exits 3 before PowerShell is even started. Both layers are tested — the validator
by fifteen unit cases, and the wiring by a case asserting that validation runs
**before** the dispatch step, so a rejected run can never reach the payload.

---

## What the tool never does

- never deletes printers, ports, or user documents outside the spool directory
- never removes printer drivers unless `-Repair -PurgeDriverStore -Force` is passed explicitly, and
  then only packages that **no** printer references
- never recurses into subdirectories
- never touches a UNC or remote file system directly
- never disables, enables or reconfigures a service permanently
- never writes credentials, GitHub tokens or personal data into logs, reports or diagnostics bundles
  (the effective-config export redacts `Update.Token`)
- never treats a **blank** selection filter as "match everything" — a blank value is refused with
  exit code 3, because in PowerShell `''` is falsy and would silently disable the filter
- never overwrites the file it is executing from — updates are staged and swapped at next start
- never installs bytes it could not verify against a published SHA-256, unless you explicitly pass
  `-SkipChecksumVerification`

---

## Verifying the tool itself

```powershell
# Authenticode
.\Clear-PrintSpoolerQueue.ps1 -VerifySignature
.\Clear-PrintSpoolerQueue.ps1 -VerifySignature -ExpectedSignerThumbprint 1A2B3C...

# Internal consistency: payload generation, guards, version parsing, exit-code
# contract, config schema, help completeness, value parsers, log location
.\Clear-PrintSpoolerQueue.ps1 -SelfTest

# Hash
Get-FileHash .\Clear-PrintSpoolerQueue.ps1 -Algorithm SHA256
```

`-SelfTest` exits non-zero if any check fails, so it belongs in a deployment pipeline: run it after
copying the files and before registering the scheduled task.

See [Security and Privacy](Security-and-Privacy).

---

## Auditing

Every run records, in the log and in the report:

who ran it (`Operator`), from which host (`HostComputer`), with which PowerShell
(`PowerShellVersion`), under which configuration (`ConfigSource`), against which targets, with
which selection filters (`Format-SelectionDescription` output is printed in the plan), which files
were deleted and which were locked, what the service did at each transition, and the final exit
code.

The JSON report carries a `schemaVersion` so a monitoring integration can detect a contract change
rather than silently mis-parsing one.

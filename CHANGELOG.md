# Changelog

All notable changes to **Print Spooler Queue Cleanup** are documented here.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the
project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [4.0.0]

A ground-up rebuild. v3.1.0 could stop the Spooler, delete `*.SHD`/`*.SPL` and restart it. That
is still the core, but it was the *only* thing the tool could do — no selection, no backup, no
diagnostics, no repair, no automation, no updates, and a thin help screen. v4 keeps the same
three-command mental model and makes it safe to run unattended on a production print server.

### Added

**Selective cleanup** — the single biggest operational gap in v3.
- `-PrinterName` / `-ExcludePrinterName` (wildcards, case-insensitive)
- `-JobId` for surgical removal of specific jobs
- `-OwnerName`, `-DocumentName` filters
- `-OlderThan` / `-NewerThan` age filters accepting `30s`, `15m`, `2h`, `7d`, `1w` or `[TimeSpan]`
- `-LargerThanBytes` / `-SmallerThanBytes` accepting `10MB`, `1.5GB` or a raw byte count
- Job inventory from `Win32_PrintJob` with a `.SHD`-header fallback for when the Spooler is too
  hung to answer WMI — i.e. precisely when the tool is needed. WMI and `.SHD` results are
  correlated so selective deletion gets both rich metadata and exact file names.

**Backup and restore** (finding A3 — v3 deleted files with no undo)
- `-Backup` archives every selected file into a ZIP with a `backup-manifest.json` and a SHA-256
- `-BackupPath`, `-BackupRequired` (abort if the backup fails), `-ArchivePath` (move bundles to a share)
- `Restore-SpoolerBackup`, `Get-SpoolerBackup`, with hash verification and zip-slip protection
- Automatic backup retention pruning (`Cleanup.BackupKeepDays`)
- The docs state plainly that restoring spool files does **not** re-queue jobs

**Diagnostics and repair** (findings A5, A6)
- `-Diagnostic`: read-only collection of OS, print-related services, `spoolsv.exe`/`splwow64.exe`
  health, printers, drivers, TCP/IP ports, spool-directory contents and age, the queued-job table,
  registry settings, and Service Control Manager + `PrintService/Admin` events
- `Test-PrintSpoolerHealth`: weighted 0–100 health score for monitoring
- `-Repair`: dependent-service cycling, `PrintNotify` recovery, `-PurgeFormsCache`,
  `-PurgeDriverStore` (unreferenced packages only, requires `-Force`), `-ResetSpoolDirectory`
- `-SelfTestPrint`: submits and immediately cancels a job to prove printing actually works
- `-DiagnosticsBundlePath`: a ZIP suitable for attaching to a support issue, with secrets redacted

**Dependent-service orchestration** (finding A1 — the most common real-world failure)
- Stops `Fax`, `PrintNotify` and `PrintWorkflowUserSvc_*` before the Spooler, restarts them after
- Only services that were actually `Running` are restarted, so a disabled Fax service stays disabled
- `Service.ForceStopFallback` escalates to terminating a service process stuck in `STOP_PENDING`
- Every state transition is recorded in `serviceTrace[]` of the report

**Robustness** (finding A4)
- Delete retries with exponential backoff (`500 ms → 1 s → 2 s`), configurable
- Locked files are counted and reported per-file instead of aborting the run
- Post-restart verification detects files the Spooler "resurrects"
- `-WaitForQuiescenceSeconds` avoids interrupting an actively draining queue
- `-Watchdog` / `-WatchdogMinutes`: an out-of-process watchdog that restarts the Spooler and
  terminates the run (exit 15) if it wedges

**Fleet execution** (finding A11)
- `-ThrottleLimit N` using PowerShell 7 `ForEach-Object -Parallel`, with an automatic runspace-pool
  fallback on Windows PowerShell 5.1
- `-PreFlightProbe` finds offline hosts in ~1 s instead of a ~40 s WinRM timeout each
- `-Credential`, `-UseSsl`, `-ComputerListFile`
- Per-computer isolation: one failure can never abort the fleet

**Observability** (findings A13, A14, A23)
- Log target chosen from the first *writable* candidate (`%ProgramData%` → `%LocalAppData%` →
  script dir → `%TEMP%`), probed with a real write rather than an ACL guess
- Size-based rotation with generations, plus age-based retention
- Pipe-delimited UTC record format, or JSON Lines with `-JsonLog`
- Levels `TRACE/DEBUG/INFO/WARN/ERROR/CRITICAL` with `-LogLevel`
- Real Windows Event Log integration (`-EventLog`, source `PrintSpoolerCleanup`) — advertised by
  the v3 README but never implemented
- `Export-RunReport` in Json / Csv / Html / Text; the HTML report is self-contained (inline CSS,
  no CDN, no script) and fully HTML-escaped
- `-MachineReadableJson`: exactly one JSON document on stdout, nothing else

**Updates** (finding A8)
- `-CheckUpdate`, `-Update`, `-AutoUpdate`, `Get-SpoolerCleanupUpdate`, `Update-SpoolerCleanup`
- Three discovery strategies: `/releases/latest` → `/tags` → raw branch file (AST-parsed version
  stamp), so discovery still works when the API is rate-limited or the repo has no releases
- SHA-256 verification against the release's `checksums.txt`; unverified bytes are refused unless
  `-SkipChecksumVerification` is passed explicitly
- **Staged** self-update via `.update\` + `pending-update.json`, applied at the next start — a
  process never overwrites the file it is executing from
- Previous version kept in `backup-<version>-<stamp>\` for rollback
- 24 h cooldown state file so scheduled runs do not hammer the API
- `PSQ_GITHUB_TOKEN` support to lift the anonymous 60 req/h limit

**Automation** (finding A7)
- `Install-SpoolerCleanupTask` / `Get-SpoolerCleanupTask` / `Start-SpoolerCleanupTask` /
  `Uninstall-SpoolerCleanupTask`
- Daily, Weekly, Hourly, AtStartup, OnEvent and Once schedules
- Always launches the `.ps1` with `-Force -NonInteractive -Quiet -NoProgress`, never the `.bat`
- Execution time limit, restart-on-failure, start-when-available; battery restrictions disabled
- Registers via the `ScheduledTasks` module or a UTF-16LE XML + `schtasks.exe` fallback

**Configuration** (finding A19)
- `PrintSpoolerCleanup.config.psd1` with auto-discovery and a documented precedence chain
  (CLI → `PSQ_*` env → `-ConfigPath` → auto-discovered → defaults)
- 26 `PSQ_*` environment variables
- Parsed with the PowerShell AST and `SafeGetValue()`: command invocations, pipelines and
  assignments are **rejected**, so a config from a share cannot execute code
- Strict schema and range validation; unknown keys abort with exit 3 rather than being ignored
- `-SaveConfig` writes the effective configuration back out

**Help** (finding A9)
- `-Help` / `-HelpSection <name>` / `Get-SpoolerCleanupHelp`: 19 sections covering usage, modes, every
  parameter, 16 worked examples, exit codes, config schema, environment variables, log and report
  formats, remote prerequisites, scheduling, update semantics, the safety model, diagnostics
  interpretation and a troubleshooting matrix
- The parameter table is generated from **live parameter metadata**, so help cannot drift from code
- `-ListHelpSections`, `-ShowVersion`, `Get-SpoolerCleanupVersion -Detailed`, `Get-ExitCodeMeaning`

**Safety guards**
- Extension allow-list (`.SPL`, `.SHD`; wildcards are rejected outright)
- Drive-letter shape validated **before** `GetFullPath()`, which is platform-relative
- Leaf must be `PRINTERS`; path must live under a `spool` folder; minimum depth; never a UNC path;
  never equal to a protected system root
- Volume guards: >100 000 files or >100 GB means "this is not a spool directory" — abort
- No recursion, ever
- `DefaultSpoolDirectory` is validated, not trusted
- `-VerifySignature` / `-ExpectedSignerThumbprint` Authenticode verification

**Engineering**
- `src/` split into `Shared`, `Private`, `Public`; `tools/Build.ps1` assembles **both** the
  standalone script and the importable module from the same sources
- `PrintSpoolerCleanup.psd1`/`.psm1` module exporting 14 cmdlets
- `tests/Run-Tests.ps1` **nine**-stage quality gate (715 checks): build → parse validation →
  23 source-lint rules → 14 batch-lint rules → 240 unit tests → 36 module-import tests →
  215 end-to-end smoke checks → 46 release-artifact integrity checks → 123 brief-compliance checks
- `tests/Unit.StrictMode.ps1` runs aggregation, serialisation and formatting under
  `Set-StrictMode -Version Latest`, with empty and heterogeneous inputs
- `tools/module-test.ps1` validates the `.psd1`/`.psm1` form, which is a genuinely separate code
  path from dot-sourcing (module scope, `Export-ModuleMember`, private-helper isolation)
- `tools/Invoke-CollisionScan.ps1` detects local variables that collide case-insensitively with a
  typed parameter, with a self-test proving the detector actually fires
- `tools/Invoke-AutomaticVariableScan.ps1` detects assignments to (and splatting of) PowerShell
  automatic variables such as `$args`, `$Matches`, `$Error` and `$_`, excluding the correct
  `$null = ...` discard idiom and the `*Preference` variables; testable via `-SourceRoots`
- `tools/Invoke-PS51CompatibilityScan.ps1` enforces the declared PowerShell **5.1 floor** locally,
  rather than only in CI and only for the remote payload. It flags `??`, `??=`, the ternary operator,
  `?.`, `clean{}`, `ConvertFrom-Json -AsHashtable`, `Join-Path -AdditionalChildPath`, and — the two
  that fail *silently* — `$IsWindows`/`$IsLinux` (which are `$null` on 5.1, so `if ($IsWindows)` takes
  the wrong branch instead of throwing) and unguarded `ForEach-Object -Parallel` (which parses on 5.1
  and then fails at bind time). Detection is token/AST based, so comments, help text, hashtable keys
  and the safe `Get-Variable -Name 'IsWindows'` probe are correctly not findings
- `tools/ps51-scanner-selftest.ps1` is its 16-fixture self-test: 10 constructs that **must** be flagged
  and 6 legitimate 5.1 patterns that **must not** be, including the guarded `-Parallel` form the
  product genuinely relies on
- `tools/Invoke-DocLint.ps1` structurally validates all 34 Markdown documents — balanced code fences,
  table column agreement, heading spacing, duplicate-heading anchor collisions, link hygiene. It found
  two real rendering defects: unescaped `|` inside table cells, which GitHub splits even inside
  backticks, so those rows rendered with an extra column
- A six-case report-schema contract suite that compares the emitted JSON **wire text** against
  `docs/Report-Schema.md` case-sensitively, in both directions, plus a case pinning the payload's job
  projection to the documented field list
- `tools/Invoke-TemplateLint.ps1` validates `.github/**/*.yml` with no YAML-parser dependency, catching
  invalid escapes in double-quoted scalars (which make GitHub render *"There was an error loading this
  form"* rather than the content) and issue-form schema violations. Block-scalar aware, so the
  PowerShell regexes embedded in `ci.yml` are not false positives. Its 12-fixture self-test asserts
  both directions
- `tools/Invoke-WorkflowLint.ps1` parse-validates the PowerShell embedded in `.github/workflows/*.yml`,
  which is invisible to every other check until it fails on a GitHub runner
- A test-harness `Invoke-WithEnvironment` helper that saves and **restores** environment variables
  instead of deleting them — a test that deleted the gate's `PSQ_LOG_PATH` silently un-redirected
  logging for every later stage
- `tools/Publish-Wiki.ps1` publishes `docs/wiki/` to the GitHub wiki repository, rewriting
  repository-file links into absolute URLs and validating every internal link first
- `tools/Verify-Release.ps1` audits the shippable artifacts: checksums, ZIP contents and
  path-traversal safety, version consistency across all five surfaces, documentation set and
  licence coherence
- `tools/Verify-Brief.ps1` proves every requirement from the original brief is still present **and
  functional** — parameter metadata, exported cmdlets, implemented update-discovery strategies,
  rendered help sections, the wiki page set, and live invocations. A refactor can silently drop a
  documented capability while every other check stays green; this stage fails the build if it does
- Dependency-free test harness (no Pester) so the suite runs on a clean machine
- `.github/workflows/ci.yml` running the gate on `windows-latest` and `ubuntu-latest`
- `checksums.txt` and generated release notes for every build
- Batch launcher rewritten as a thin delegating front-end with engine detection (PS7 preferred),
  UAC self-elevation, staged-update swap, correct non-interactive detection and full CLI parity

### Changed

- **BREAKING**: exit-code contract extended and renumbered. v3 used `0/1/2/3/4`;
  v4 uses `0–15` (see [Exit Codes](./docs/Exit-Codes.md)). `2` now means *Failed* and `1` means
  *Partial success*, which is the reverse of v3's `1 = fatal` / `2 = warnings`.
- Minimum service timeout raised from 30 s to 60 s (a large queue on slow storage needs it).
- The batch launcher no longer implements its own cleanup; it delegates to the `.ps1`. A standalone
  batch mode is intentionally not provided — one destructive implementation, not two.
- README license badge corrected from "Public Domain" to GPL-3.0-or-later to match `LICENSE`
  (finding A15). Version references no longer drift between README, script and launcher (A16).
- Unicode glyphs are now emitted only when the console code page can render them; `-Ascii` forces
  the ASCII palette, and redirection disables colour automatically.
- `$global:LASTEXITCODE` assignment replaced by a documented epilogue that publishes the exit code
  correctly for `-File`, direct invocation and dot-sourcing (finding A21).

### Fixed

- **Bare `-Help` - the exact command the README advertises - did not bind.**
  `-Help` was declared as a `[string]` section selector, so invoking it without a
  value failed with "Missing an argument for parameter 'Help'". Every in-process
  test passed a value and the batch launcher served `/help` locally, so the whole
  gate stayed green while the primary documented invocation was broken. `-Help` is
  now a switch that renders the complete reference; the section selector moved to
  `-HelpSection <name>`. The gate now runs the documented CLI forms byte-for-byte
  in a child process against the shipped artifact (defect 46).
- **Every malformed invocation exited 0.** The standalone entry point's `catch` -
  the handler for failures before the function body runs, such as parameter binding
  errors - assigned `$global:LASTEXITCODE = 3`, but the next line re-read
  `RuntimeState.ExitCode` (still "Success", because the body never ran) and
  overwrote it with 0. Schedulers and RMMs saw every bad command line as a
  success. The catch now owns the final code: binding failures exit 3
  (BadArguments) per the exit-code contract (defect 47).
- **The full reference rendered as silent empty output under
  `Set-StrictMode -Version Latest`.** The Examples section read its optional note
  key as a property (`$e.N`); five of the fifteen examples have no such key, which
  strict mode turns into a thrown "property not found". The product's own handler
  swallowed it to the log and exited 2, so on hardened hosts `-Help` produced
  nothing while every individual section still rendered and every lax-mode test
  passed. Optional hashtable keys are now read through the indexer, and the strict
  unit suite renders the full document (defect 48 - the same class as the
  StrictMode crash fixed below).
- **The test suite's verdict depended on how it was invoked.** `tools/Build.ps1`
  sets `Set-StrictMode -Version 2.0` and then runs `tests/Run-Tests.ps1` with `&`,
  which is the *same process*, so the strict mode leaked into the entire suite.
  Run directly the suite was lax and green; run through the build it was strict and
  red — on identical code. A result that changes with the invocation path is not a
  result. The gate now pins `Set-StrictMode -Version Latest` itself, so both paths
  are deterministic and the strictest setting is the default rather than an
  accident of who called it.
- **Two more documented `PSQ_*` variables were dead** — the same class as the
  `PSQ_LOG_PATH` defect above, found by auditing the *entire* advertised surface
  rather than the one variable that happened to fail. `PSQ_NONINTERACTIVE` was the
  dangerous one: an operator who set it for a scheduled task or a CI job still got
  an interactive confirmation prompt, and with no console attached the run **hung
  until the task timeout** instead of failing fast. `PSQ_CONFIG_PATH` meant a
  machine-wide policy file deployed through Group Policy was silently ignored and
  the run used built-in defaults. Both are now honoured, with precedence
  `-Parameter` > `PSQ_*` > auto-discovery, and an environment-variable contract
  suite audits all 26 advertised variables in both directions so neither a dead
  nor an undocumented one can return.
- **The JSON report did not match its own normative schema.** `docs/Report-Schema.md`
  is marked normative and ships `jq` and Python consumption examples, but
  `ConvertTo-JsonReport` serialised the PowerShell object directly and emitted
  `PascalCase` keys inside a `camelCase` envelope. All **70** documented field names
  therefore resolved to nothing for any case-sensitive consumer — jq, Python, and
  every RMM parser. It was invisible from PowerShell, where `$o.summary.Tool` and
  `$o.summary.tool` both work, so the existing assertion on
  `$obj.summary.Results[0].ComputerName` passed while the published contract was
  broken. Keys are now projected through one mechanical camelCase rule
  (`ConvertTo-CamelCaseDeep`), values are never renamed, and a six-case contract
  suite compares the emitted **wire text** against the schema document in both
  directions.
- **The payload silently dropped two documented job fields.** `PagesPrinted` and
  `Priority` are computed on both the WMI path and the `.SHD` fallback and both
  appear in the `jobs[]` schema table, but the final `Select-Object` projection
  listed every field except those two — so the schema promised columns no consumer
  would ever receive. The projection is now compared against the schema document
  rather than a hand-copied list, so a newly documented field cannot be dropped
  silently.
- **`.github/ISSUE_TEMPLATE/config.yml` did not parse at all**, so the repository's
  issue chooser was broken. A Windows path inside a *double*-quoted YAML scalar
  (`".\Clear-PrintSpoolerQueue.ps1 -Help"`) makes `\C` an unknown escape, which is
  a hard parse error. GitHub does not fail a build over it — it renders *"There was
  an error loading this form"*, so nobody could file a bug report and nothing in
  the suite noticed. Now single-quoted, and lint-guarded.
- **A blank selection filter silently widened the blast radius — the most
  dangerous defect found in this work.** An empty string is *falsy* in PowerShell,
  and the cleanup payload gates every filter on a truthiness test:

  ```powershell
  if ($Options['PrinterName'] -and @($Options['PrinterName']).Count -gt 0)
  ```

  So `-PrinterName ''` applied **no filter at all**: the run cleaned **every**
  printer and every job instead of the one the operator named. The mirror case was
  worse — a blank `-ExcludePrinterName` silently withdrew a protection that had
  been explicitly requested. This was reachable from the `.bat` launcher by typing
  `/printer` and forgetting the value, which forwarded `-PrinterName ""`.

  Blank values on the thirteen selection and path parameters are now **rejected**
  with exit code 3 and a message that explains the consequence rather than merely
  refusing. The launcher also refuses a value-taking switch given no value, so the
  mistake is caught before PowerShell is even started. Pinned by 15 unit tests,
  including one that asserts the validator is wired in *before* the dispatch step.
- **`New-ChildArgumentList` built its result in `$args`** — a PowerShell
  *automatic* variable holding the current invocation's unbound arguments. It
  worked only because the helper is always called with named parameters; any call
  passing a positional argument would have had it silently swallowed, and the
  returned vector would have contained the **caller's** arguments. This helper
  builds the command line for `-SelfElevate`, so the failure mode was a child
  process launched with the wrong arguments. Renamed to `$childArgs`. The same
  pattern had also shipped in the CI release job (`gh @args`).
  `tools/Invoke-AutomaticVariableScan.ps1` now fails the build on either form,
  with self-tests in both directions.
- **CI would have failed on a clean checkout.** The gate's build stage ran with
  `-NoZip`, but the release-integrity stage validates the ZIP: its contents, CRC,
  path-traversal safety and sync with the sources. On a fresh clone the stage
  threw an unhandled `MethodInvocationException` opening an archive that did not
  exist. It passed locally only because a stale ZIP from an earlier full build was
  still sitting in `dist/`. The gate now builds the archive it intends to verify,
  and `Verify-Release.ps1` **reports** a missing ZIP instead of crashing on it.
- **The whole product crashed under `Set-StrictMode -Version Latest`.** Two
  distinct shapes, both invisible in lax mode:
  `Measure-Object` returns `$null` (not an object) for an **empty** pipeline, so
  `.Sum` threw *"The property 'Sum' cannot be found on this object"*; and reading a
  property that a particular result object does not carry (`$_.BackupPath` on a
  `Diagnostic` result) threw the same way. Nine sites summed with `Measure-Object`
  and six read properties directly. An empty spool directory is the **normal,
  healthy** case, so this was not an edge case — any hardened host would have
  failed on the most common run. All aggregation now goes through
  `Get-ResultPropertyValue`, which is correct in every strict-mode setting and for
  every collection shape. A dedicated `tests/Unit.StrictMode.ps1` suite (13 cases)
  and a source-lint rule keep the class out.
- **`PSQ_LOG_PATH` and `Logging.Path` were completely ignored.** The log target was
  resolved from the `-LogPath` *parameter* alone, in the `begin` block, before the
  configuration was ever loaded. Both the documented environment variable and the
  config-file setting were dead, so a site that configured a central log location
  silently got auto-discovery next to the script instead. Configuration is now
  loaded in the `begin` block and the precedence is
  `-LogPath` > `PSQ_LOG_PATH` > `Logging.Path` > auto-discovery.
- **`-SelfTest` littered the release folder.** Its log-location check called
  `Resolve-LogTarget` with no explicit path, and that resolver *creates* each
  candidate directory as it probes it - including `<script dir>\Logs`. Running
  `-SelfTest` from a `dist/` folder therefore polluted the shipped artifacts. The
  check is now hermetic.
- **`-SelfTest -PassThru` (and every other `-PassThru` mode) threw when the tool
  was loaded as a module.** The entry point assigned its result to `$selfTest`,
  which in PowerShell is the *same variable* as the `[switch]$SelfTest` parameter
  (variables are case-insensitive), so a `PSCustomObject` was assigned to a
  `SwitchParameter`: *"Cannot convert value … to type SwitchParameter"*. It did
  not reproduce when dot-sourcing the script, which is why 400+ checks missed it
  until a dedicated module-import stage was added to the quality gate. The local
  is now `$selfTestResult`, and `tools/Invoke-CollisionScan.ps1` plus two lint
  rules prevent the whole class from returning.
- **Update cooldown never worked.** `Read-UpdateState` enumerated `$state.Keys` while assigning
  `$state[$k]`, which throws *"Collection was modified"* mid-loop; the surrounding empty `catch {}`
  swallowed it and returned an empty state, so every run re-checked GitHub. Keys are now snapshotted
  and the catch logs.
- `-OlderThan 30` (a bare number) was interpreted as **30 days** by `[TimeSpan]::TryParse` instead
  of 30 seconds. Bare numbers are now handled explicitly before the fallback parser.
- `switch -regex` in the time-span parser matched multiple clauses (`'ms'` satisfied both the
  millisecond and minute patterns) and returned an **array** of two `TimeSpan` values. Patterns are
  now fully anchored and the result is forced to a scalar.
- `return switch ...` is a syntax error in PowerShell (switch is a statement, not an expression).
- `'^m$|^min'` — a `$` immediately before a closing single quote forms the special variable `$'`
  and silently swallows the string terminator, breaking the parse hundreds of lines later.
- `.ContainsKey()` was called on `[ordered]@{}` dictionaries, which are
  `OrderedDictionary` and only expose `.Contains()` — every logging call threw.
- Log retention only matched one file-naming scheme, so session-stamped logs were never pruned and
  grew without bound.
- `Format-ByteSize` dropped the sign of negative values; `Format-Duration` could not accept `$null`
  because a `[TimeSpan]`-typed parameter rejects it before the null check runs.
- `Test-IsLocalTarget` trimmed `'.'` to an empty string via `TrimEnd('.')`, so the `.` alias for
  "this machine" was never recognised.
- `New-ChildArgumentList` emitted `-Name:` with no value for booleans (`"$$:"` parses as the
  automatic variable `$$` followed by a colon), breaking self-elevation.
- `PSQ_PRINTER_NAME` used `.Split(',', ';')`, which binds to `Split(char, int)` and truncated the
  value instead of splitting on both delimiters.
- `Test-Configuration` treated an empty hashtable as valid (`-not @{}` is `$false` in PowerShell).
- The spool-directory guard accepted `C:\Users\PRINTERS`: the forbidden-root check compared for
  equality only. A structural "must live under a `spool` folder" rule now defeats that whole class
  of corrupt-registry accident.
- `ConvertTo-SafeFileName` relied on `[IO.Path]::GetInvalidFileNameChars()`, which reflects the
  *host* platform and omits `\ : "` on Linux — the tool only ever runs on Windows.
- The log writability probe wrote a blank line into every log file, breaking strict parsers.
- Running the tool from a read-only directory (Program Files, a locked share) disabled logging
  entirely instead of falling back to a writable location.
- 45 silent `catch { }` blocks now carry an explicit rationale comment, and two lint rules prevent
  both that and mutation-during-enumeration from returning.
- The v3 README's `schtasks` example ran the interactive `.bat` as SYSTEM, so "Press any key"
  blocked the task until its time limit killed it.

### Removed

- The batch file's own stop/delete/start implementation (two destructive code paths for one job).
- The misleading "Public Domain" license badge.

---

## [3.1.0]

- Fixed invalid top-level `begin`/`process`/`end` blocks caused by `Set-StrictMode` between
  `param()` and the named blocks.
- Routed all remote service operations through `Invoke-Command` (`Get-Service`/`Stop-Service`
  `-ComputerName` were removed in PowerShell 7).
- Fixed `Clear-SpoolFiles` referencing an undefined `$ScriptConfig.SpoolDirectory` remotely.
- Restricted cleanup to `*.SHD`/`*.SPL`.
- Replaced `$Global:Statistics` with `$script:Stats`; added per-computer success/failure tracking.
- Batch: locale-independent size accounting, correct exit codes, `/WhatIf`, `/Force`, `/Quiet`,
  `/NoPause`, `/Log`, `/UsePS`, `DefaultSpoolDirectory` resolution, best-effort recovery.

## [3.0.0] and earlier

- Initial PowerShell and batch implementations of stop → delete → restart.

[4.0.0]: https://github.com/paulmann/Print-Spooler-Queue-Cleanup/releases/tag/v4.0.0
[3.1.0]: https://github.com/paulmann/Print-Spooler-Queue-Cleanup/releases/tag/v3.1.0

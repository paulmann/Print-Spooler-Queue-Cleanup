# Handover Report — Print Spooler Queue Cleanup v4.0.0

**Baseline:** `paulmann/Print-Spooler-Queue-Cleanup` @ `a70a430` (v3.1.0 — 518-line `.ps1`,
472-line `.bat`, 822-line README)
**Result:** a production-ready v4.0.0 with a build pipeline, a **715-check nine-stage quality gate**
(verified stable across three consecutive runs, and re-verified from a clean checkout on
Linux / PowerShell 7.6.6, 2026-10-10), and a 25-page English wiki.

---

## 1. Delivery summary

| Metric | Value |
|---|---|
| Source under `src/` | 10 968 lines across 18 files |
| Batch launcher template | 543 lines |
| Tests | 251 unit + 36 module + 215 smoke + 47 artifact-integrity + 129 brief-compliance |
| Build/validation tooling | **15 scripts in `tools/`**: `Build`, `Invoke-Validation`, `Invoke-CollisionScan`, `Invoke-AutomaticVariableScan`, `Invoke-PS51CompatibilityScan`, `ps51-scanner-selftest`, `Invoke-WorkflowLint`, `Invoke-DocLint`, `Invoke-TemplateLint`, `template-lint-selftest`, `module-test`, `smoke-test`, `Verify-Release`, `Verify-Brief`, `Publish-Wiki` |
| Wiki | 25 files — 23 content pages plus `Home.md` and `_Sidebar.md` |
| Engineering docs | 3 (exit-code contract, report schema, refactor plan) |
| Governance | README, CHANGELOG, SECURITY, CONTRIBUTING, CODE_OF_CONDUCT, issue templates, CI |
| Quality gate | **9/9 stages · 715 checks · PASSED** (stable, incl. the nested Build→gate path; re-verified on a clean checkout, 2026-10-10) |
| Built artifacts | standalone `.ps1` (524 KB), launcher `.bat` (30 KB), module `.psd1`/`.psm1`, ZIP (302 KB), `checksums.txt`, `RELEASE-NOTES.md` |

Three consumable forms, all generated from one source tree so they cannot drift:

```
dist/Clear-PrintSpoolerQueue.ps1                    standalone script
dist/Clear-PrintSpoolerQueue.bat                    double-click launcher
dist/PrintSpoolerCleanup/PrintSpoolerCleanup.psd1   importable module (14 cmdlets)
```

---

## 2. What was wrong with v3.1.0 (23 findings)

The full audit is in [`REFACTOR-PLAN.md`](./REFACTOR-PLAN.md). The material ones:

| # | Finding | Severity |
|---|---|---|
| A1 | `Spooler` stopped without stopping dependents → stop hangs on Win11/Server | High |
| A2 | No selective cleanup — the whole queue was always destroyed | High |
| A3 | No backup, no undo | High |
| A4 | Locked files: single attempt, no retry, no resurrection detection | High |
| A5 | No diagnostics — no way to answer *why* the queue is stuck | High |
| A6 | No repair path for the actual root causes | High |
| A7 | No scheduled/automated mode; the README's `schtasks` example hung forever | High |
| A8 | No update mechanism | Med |
| A9 | Thin help; no exit-code/env/config reference | Med |
| A11 | Sequential remote execution only (200 hosts × ~15 s) | Med |
| A14 | Log written next to the script → silently lost in Program Files / read-only shares | Med |
| A15 | README claimed "Public Domain" while `LICENSE` is GPL-3.0 | Med |
| A16 | Version drift: README said "v3.0 Professional", scripts said 3.1.0 | Med |
| A21 | `$global:LASTEXITCODE` assignment does not set the process exit code for `-File` | Low |
| A22 | No non-interactive guard → hangs under Task Scheduler | Low |
| A23 | README advertised Event Log integration that was never implemented | Low |

---

## 3. What was added

**Precision** — `-PrinterName`, `-ExcludePrinterName`, `-JobId`, `-OwnerName`, `-DocumentName`,
`-OlderThan`, `-NewerThan`, `-LargerThanBytes`, `-SmallerThanBytes`. Inventory from
`Win32_PrintJob` correlated with `.SHD` headers so selective deletion gets both rich metadata and
exact file names; jobs with unreadable metadata are **not** guessed at.

**Safety** — `-Backup`/`-BackupRequired`/`-ArchivePath` with SHA-256 manifests,
`Restore-SpoolerBackup`, exponential-backoff retries, post-restart resurrection detection, four
independent delete guards, a validated (not trusted) registry override.

**Root cause** — `-Diagnostic` (10 evidence areas), `Test-PrintSpoolerHealth` (0–100 score),
`-Repair`, `-PurgeFormsCache`, `-PurgeDriverStore`, `-ResetSpoolDirectory`, `-SelfTestPrint`,
`-Watchdog`, `-DiagnosticsBundlePath`.

**Scale** — `-ThrottleLimit` (PS7 `-Parallel`, PS5.1 runspace pool fallback), `-PreFlightProbe`,
`-Credential`, `-UseSsl`, `-ComputerListFile`, per-computer isolation.

**Observability** — rotating + retention-pruned logs in the first *writable* location, pipe-delimited
UTC or JSON Lines, 6 levels, real Event Log integration, JSON/CSV/HTML/Text export,
`-MachineReadableJson` (exactly one JSON document on stdout), HTML-escaped self-contained reports.

**Lifecycle** — `-CheckUpdate`/`-Update`/`-AutoUpdate` with three discovery strategies, SHA-256
verification, staged self-update, cooldown state; `Install-SpoolerCleanupTask` and friends.

**Ergonomics** — `-Help` / `-HelpSection` (19 sections, generated from live parameter metadata), 26 `PSQ_*`
environment variables, AST-sandboxed config file with strict validation, `-SelfElevate`,
`-VerifySignature`, `-Ascii`/`-NoColor` auto-detection, `-PassThru`, `-SelfTest`.

**Engineering** — `src/` → `tools/Build.ps1` → both distributables; deterministic manifest;
nine-stage, 697-check quality gate; CI on Windows **and** Linux; `checksums.txt`; generated
release notes.

---

## 4. Defects found and fixed *during* this work

These are worth reading, because each one was invisible in review and each was caught by a test or by
running on the "wrong" platform.

**This table is the single canonical defect register.** Numbers 1–20 were found in the first pass;
21–45 in the completion sweep described in §6b; 46–48 in the release-verification pass described
in §6c. Source comments, CI comments and the wiki cite these
numbers, so a number means the same thing everywhere — §6b deliberately carries no table of its own.

| # | Defect | How it was caught |
|---|---|---|
| 1 | **`Read-UpdateState` enumerated `$state.Keys` while assigning `$state[$k]`** → "Collection was modified", swallowed by an empty `catch {}`, returning an empty state. **The update cooldown never worked**, so every run hit the GitHub API. | Unit test on state round-trip |
| 2 | `return switch` is a **syntax error** — `switch` is a statement, not an expression | Parser gate |
| 3 | `switch -regex` does not stop at the first match and is a substring match: `'ms'` satisfied both the millisecond and minute clauses, returning an **array of two TimeSpans** | Unit test |
| 4 | `'^m$\|^min'` — a `$` before a closing single quote forms the special variable **`$'`** and swallows the terminator; the parse error surfaced 100 lines away | Parser gate |
| 5 | `.ContainsKey()` on `[ordered]@{}` — an `OrderedDictionary` has only `.Contains()`. **Every logging call threw.** | Smoke test |
| 6 | PowerShell variables are **case-insensitive**: closures `$W`/`$P`/`$Sub` were silently overwritten by loop variables `$w`/`$p`/`$section` | Smoke test |
| 7 | A loop variable `$section` collided with a `[ValidateSet()]` `$Section` **parameter**, so assigning `'Cleanup'` threw a validation error | Unit test |
| 8 | `-OlderThan 30` was read as **30 days** by `[TimeSpan]::TryParse` instead of 30 seconds | Unit test |
| 9 | `GetFullPath('C:\PRINTERS')` is **platform-relative** — on Linux it returned `<cwd>/C:\PRINTERS`, silently satisfying the depth guard | **Linux CI leg** |
| 10 | The spool guard accepted `C:\Users\PRINTERS` (right leaf, right depth, not equal to a protected root) → added the structural "must live under a `spool` folder" rule | Unit test |
| 11 | `GetInvalidFileNameChars()` reflects the **host** platform, omitting `\ : "` on Linux | Unit test |
| 12 | `.Split(',', ';')` binds to `Split(char, int)` and **truncates** instead of splitting on both | Unit test |
| 13 | `"$$:"` in a double-quoted string parses as the automatic variable `$$` → `-Name:` with no value, breaking self-elevation | Smoke test |
| 14 | `Test-IsLocalTarget` trimmed `'.'` to empty via `TrimEnd('.')`, so the `.` alias never matched | Smoke test |
| 15 | Log retention matched only one naming scheme → session-stamped logs were **never pruned** | Unit test |
| 16 | The log writability probe wrote a blank line into every log, breaking strict parsers | Unit test |
| 17 | `Test-Configuration` treated `@{}` as valid (`-not @{}` is `$false`) | Unit test |
| 18 | `Initialize-RuntimeState` read `$Script:RuntimeState` before assignment → **throws under `Set-StrictMode`** at import time | Build-time run |
| 19 | 45 silent `catch { }` blocks — the mechanism that hid defect #1 | New lint rule |
| 20 | An unassigned `Resolve-PayloadScriptBlock -Force` emitted a scriptblock into the pipeline, which bound to the **next command's** parameters | Smoke test |
| 21 | **`$selfTest` (local) is the same variable as `[switch]$SelfTest` (parameter)** — assigning the result object threw *"Cannot convert value … to SwitchParameter"*. Reproduced **only via `Import-Module`**, not when dot-sourced, so 400+ checks stayed green while `-PassThru` was broken for every module user | New module-import stage (6) |
| 22 | `Test-ModuleManifest.ExportedVariables` is a `Dictionary`, so `@($dict).Count` is `1`, not `0` — the *assertion* was wrong, not the manifest | Count `.Keys` |
| 23 | Exercising the module wrote `Logs/` **inside the release artifact folder** | Redirect logging via `PSQ_LOG_PATH`; assert the tree stays clean |
| 24 | `Join-Path` throws `Cannot find drive 'C'` on a host with no C: drive — a noisy false alarm inside `-SelfTest` | Plain concatenation |
| 25 | **`PSQ_LOG_PATH` and `Logging.Path` were entirely dead.** The log target was resolved from the `-LogPath` parameter alone, in `begin{}`, before configuration was loaded — so a documented site-wide setting silently did nothing and every run fell back to auto-discovery | Release-integrity stage (8): a stray `Logs/` appeared in `dist/` |
| 26 | `-SelfTest` created `<script dir>\Logs`, because `Resolve-LogTarget` **creates** each candidate directory while probing it — so a self-test polluted the release folder | Release-integrity stage (8) |
| 27 | Two lint assertions read the global `$LASTEXITCODE`, which any earlier external command can leave behind — **flaky** when the gate ran nested inside `Build.ps1` | Child process; read `Process.ExitCode` |
| 28 | A test did `Remove-Item Env:\PSQ_LOG_PATH`, deleting the **gate's** redirection and un-redirecting logging for every later stage — presenting as three failures in unrelated assertions | `Invoke-WithEnvironment` save/restore helper |
| 29 | **The product crashed under `Set-StrictMode -Version Latest`.** `Measure-Object` returns `$null` for an *empty* pipeline, so `.Sum` threw (9 sites); reading a property absent from a heterogeneous result threw too (6 sites). An empty spool directory is the **normal** case, so any hardened host would have failed on the most common possible run | New StrictMode suite (13 cases) |
| 30 | `Resolve-LogTarget -ExplicitPath $null` in a unit test created `dist/Logs` — a test with side effects on the artifact tree | Hermetic scratch `ScriptDirectory` |
| 31 | The collision detector had **no self-test** — a lint rule that cannot fail is not a lint rule | Synthetic fixture; asserts it fires *and* names the parameter |
| 32 | The collision scanner **over-reported**: it collected nested scriptblock param blocks (e.g. the `ForEach-Object -Parallel` worker) as if they were the function's own | Only the function's own `ParamBlock`; alias-aware casts |
| 33 | **`New-ChildArgumentList` built its result in `$args`** — the automatic variable holding the invocation's unbound arguments. It worked only because the helper is always called with named parameters; a positional call would have returned the *caller's* arguments. This helper builds the `-SelfElevate` command line | New automatic-variable scanner |
| 34 | **CI would have failed on a clean checkout.** The gate built with `-NoZip` while stage 8 validated the ZIP, throwing an unhandled `MethodInvocationException`. It passed locally only because a *stale* ZIP from an earlier build was still in `dist/` | Reproduced by deleting `dist/` first |
| 35 | The CI release job splatted `gh @args` — the same automatic-variable collision as 33, in the publishing path | New automatic-variable scanner + workflow linter |
| 36 | `Verify-Brief` asserted a help section named `Synopsis` (the product publishes `Usage`) and demanded diagnostic evidence areas off-Windows, where the cmdlet correctly **refuses** | Assert the real names; make the check platform-aware |
| 37 | **Two Markdown tables rendered with a spurious extra column on GitHub.** A literal `\|` inside a table cell is treated as a cell separator by GFM *even inside backticks*, so a row documenting the regex `'^m$\|^min'` and a row containing the pipeline `Get-SpoolerCleanupVersion \| Select-Object …` each gained a column. No functional test can see this: Markdown has no compiler | New `Invoke-DocLint` stage |
| 38 | The same defect was **re-introduced while writing the row that documents the linter which catches it** — the description contained a literal `\|` inside backticks. The gate failed on the very next run, which is the strongest available evidence that the check works | `Invoke-DocLint` (self-inflicted, caught immediately) |
| 39 | **A blank selection filter silently widened the blast radius.** An empty string is *falsy* in PowerShell, and the payload gates every filter on `if ($Options['PrinterName'] -and @(...).Count -gt 0)` — so `-PrinterName ''` applied **no filter at all** and the run cleaned **every** printer instead of the one named. The mirror case was worse: a blank `-ExcludePrinterName` silently withdrew a protection the operator had asked for. Reachable from the launcher by typing `/printer` and forgetting the value, which forwards `-PrinterName ""`. | New launcher-semantics review; fixed in the validator, the launcher **and** pinned by 15 unit tests |
| 40 | **`.github/ISSUE_TEMPLATE/config.yml` did not parse at all.** A Windows path inside a DOUBLE-quoted YAML scalar (`".\Clear-PrintSpoolerQueue.ps1 -Help"`) makes `\C` an unknown escape, which is a hard YAML parse error. GitHub's response is not an error at build time — it renders *"There was an error loading this form"*, so nobody could file a bug report and no check in the suite noticed | New `Invoke-TemplateLint` stage |
| 41 | `Invoke-TemplateLint` **could not fail**: it resolved the tree from its own location, so the self-test's `-WorkingDirectory` fixture was ignored and all eight known-broken templates "passed". Caught only because the self-test asserts the detector *fires*, not merely that it runs | Template-lint self-test (12 fixtures, both directions) |
| 42 | **The JSON report did not match its own normative schema.** `docs/Report-Schema.md` documents `camelCase` keys and ships `jq` and Python examples, but `ConvertTo-JsonReport` serialised the PowerShell object directly, emitting `PascalCase` inside a `camelCase` envelope. All **70** documented field names therefore resolved to nothing for any case-sensitive consumer (jq, Python, RMM). Invisible from PowerShell, where `$o.summary.Tool` and `.tool` both work — so the existing test asserting `$obj.summary.Results[0].ComputerName` passed while the contract was broken | New report-schema contract suite (6 cases, matched against the raw wire text) |
| 43 | **The payload silently dropped two documented job fields.** `PagesPrinted` and `Priority` are computed on both the WMI path and the `.SHD` fallback, and both are in the `jobs[]` schema table — but the final `Select-Object` projection listed every field except those two, so the schema promised columns no consumer ever received | Projection now compared against the schema document, not a hand-copied list |
| 44 | **Two more documented `PSQ_*` variables were dead** — the same class as defect 25, found by auditing the *whole* advertised surface instead of one variable. `PSQ_NONINTERACTIVE` was the dangerous one: an operator who set it for a scheduled task or CI job still got an interactive confirmation prompt, and with no console attached the run **hung until the task timeout** instead of failing fast. `PSQ_CONFIG_PATH` meant a machine-wide policy file set through Group Policy was silently ignored and the run used built-in defaults | New environment-variable contract suite (audits all 26 in both directions) |
| 45 | **The test suite's outcome depended on how it was invoked.** `Build.ps1` sets `Set-StrictMode -Version 2.0` and then calls `Run-Tests.ps1` with `&` — the *same process* — so the strict mode leaked into the whole suite. Run directly the suite was lax and green; run through the build it was strict and red. A result that changes with the invocation path is not a result. The gate now pins `Set-StrictMode -Version Latest` itself, so both paths are deterministic, and the strictest setting is the default | Comparing the direct and nested runs line by line |
| 46 | **Bare `-Help` - the exact command the README quick-start quotes - did not bind.** `-Help` was declared `[string]` (a section selector), so invoking it with no value threw "Missing an argument for parameter 'Help'". Every in-process test passed a section value and the launcher served `/help` locally, so 697 checks stayed green while the primary documented invocation was broken. | New CLI-contract suite: the documented forms invoked byte-for-byte in a child `pwsh -File` process against the shipped artifact |
| 47 | **The entry point's binding-error catch overwrote its own exit code.** It set `$global:LASTEXITCODE = 3`, then the next line re-read `RuntimeState.ExitCode` - still "Success", because the function body never ran - and exited **0**. Every malformed invocation reported success to schedulers and RMMs. | The same child-process CLI contract: an invalid argument must exit 3 (BadArguments), never 0 |
| 48 | **The full reference rendered as silent empty output under `Set-StrictMode -Version Latest`.** The Examples section read its optional note key as a property (`$e.N`); 5 of 15 entries lack that key, strict mode throws "The property 'N' cannot be found", and the product's own handler swallowed it to the log and exited 2. Hardened hosts got nothing from the tool's own manual while every single section still rendered and every lax-mode test passed. | The strict unit suite now renders the FULL document end-to-end, not just aggregation and formatting |

Every one now has a regression test, and the recurring classes (2, 4, 5, 6, 19, 21) have **source-lint
rules** in the gate so they cannot come back.

Defect 21 is the most instructive. It is invisible to a parser, invisible to a type checker, and
invisible to any test that loads the product the way the developer loads it. It was found only because
a **separate gate stage exercises the second shipping form**. The lesson is recorded in the wiki as
trap 4(c), and `tools/Invoke-CollisionScan.ps1` now fails the build on any local variable that
collides case-insensitively with a typed parameter — with a self-test proving the detector fires.

---

## 5. Quality gate

`./tests/Run-Tests.ps1` — six stages, exit 0 only if all pass:

| Stage | Result |
|---|---|
| 1. Build | both distributables assembled from `src/` |
| 2. Parse validation | 36 files + artifacts, **0 errors** |
| 3. Source lint | **23/23** |
| 4. Batch lint | **14/14** (labels, shifts, duplicates, ASCII, CRLF, no own delete loop) |
| 5. Unit tests | **251/251** (incl. the 13-case StrictMode suite and the 11-case Help CLI-contract suite) |
| 6. Module import | **36/36** (manifest, export surface, isolation, artifact hygiene) |
| 7. End-to-end smoke | **215/215** |
| 8. Release artifact integrity | **47/47** (checksums, ZIP safety, version coherence, docs, licence) |
| 9. Brief compliance | **129/129** (every requirement from the original brief, exercised, incl. child-process CLI contracts) |
| **Total** | **715 checks · PASSED** |

Design choices:

- **No Pester.** The product promises "no external modules"; the suite holds itself to the same rule
  so it runs on a clean machine and on a Linux runner.
- **Unit tests run against `dist/`**, not `src/`, so a build defect cannot hide behind a source-level
  pass.
- **The Linux leg is load-bearing.** It caught defect #9 — a safety guard that silently passed on
  Windows and was meaningless elsewhere.
- **Live network test** against the real GitHub API, which validated all three discovery strategies:
  the repo has no releases and no tags, so discovery correctly fell through to the raw-file strategy
  and read `3.1.0`.

---

## 6. Documentation

**Wiki** (`docs/wiki/`, 23 content pages + Home + `_Sidebar`, mirrors the GitHub wiki):
Home, `_Sidebar`, Getting Started, Installation and Deployment, Batch Launcher Reference,
CLI Reference, Cmdlet Reference, Configuration, Exit Codes, Logging and Event Log,
Reports and Machine-Readable Output, Remote and Fleet Management, Scheduling and Automation,
Updates and Versioning, Enterprise Deployment, Troubleshooting, FAQ, Print Spooler Internals,
Why Print Queues Get Stuck, Safety Model, Security and Privacy, For Developers.

**Engineering docs** (`docs/`): `Exit-Codes.md` (normative contract), `Report-Schema.md` (normative
JSON/CSV/HTML schema with versioning rules), `REFACTOR-PLAN.md` (the 23-finding audit).

**Governance:** README (license badge corrected to match `LICENSE`), CHANGELOG (Keep a Changelog,
with every fix attributed), SECURITY (scope, hardening guidance, disclosed limitations), CONTRIBUTING
(the 15 PowerShell traps, what gets rejected and why), CODE_OF_CONDUCT, structured issue templates.

The built-in `-Help` renders the same material **from live parameter metadata**, so the console
reference cannot drift from the code — `-SelfTest` fails if a parameter lacks a description.

---

## 6b. Completion sweep (second pass)

A completion audit after the first "100%" found **twenty-five more real defects**. All are fixed and
regression-tested, and each one produced a new permanent check rather than just a patch.

The twenty-five defects are listed in the canonical register in **§4** (numbers 21–45) rather than
duplicated here. Earlier revisions of this document kept a second, competing numbering in this
section, and the two had already drifted — §4 used 28 for the `ExportedVariables` assertion while this
section used 28 for the `PSQ_LOG_PATH` test, and 22/27 appeared twice with different meanings. A
defect register that disagrees with itself is worse than none, because source comments cite it. One
table, one numbering.

Of the twenty-five, **eleven were product bugs** (21, 25, 26, 29, 33, 34, 35, 39, 42, 43, 44), **eleven
were test-harness bugs** (22, 23, 24, 27, 28, 30, 31, 32, 36, 41, 45), **two were documentation-rendering
bugs** (37, 38) and **one was a repository-metadata bug** (40).

Defect 29 is the most serious. It is invisible in the development host, invisible to the parser, and
invisible to every test written in lax mode — yet it would have broken the tool for any customer who
hardened PowerShell, on their *first, cleanest* run.

Defect 34 is the most instructive about process: the suite was green, the artifacts looked correct,
and the failure appeared only on a clean checkout, because a stale ZIP from an earlier build was
silently satisfying a check that should have found nothing to check.

### Tooling added

| Tool | Purpose |
|---|---|
| `tools/module-test.ps1` | Stage 6 — validates the `.psd1`/`.psm1` form: manifest, export surface, private-helper isolation, artifact hygiene |
| `tools/Invoke-CollisionScan.ps1` | AST scanner for local/parameter name collisions; self-tested |
| `tools/Verify-Release.ps1` | Stage 8 — checksums, ZIP contents and path-traversal safety, version coherence across all five surfaces, documentation set, licence coherence |
| `tools/Verify-Brief.ps1` | Stage 9 — proves every requirement from the original brief is still present *and functional*: parameter metadata, exported cmdlets, implemented update-discovery strategies, rendered help sections, the wiki page set, live invocations. Platform-aware (asserts the diagnostic cmdlet **refuses** off Windows rather than inventing evidence) |
| `tools/Invoke-AutomaticVariableScan.ps1` | Detects assignment to (or splatting of) PowerShell automatic variables — `$args`, `$Matches`, `$Error`, `$_`. Excludes the correct `$null = ...` discard idiom and the `*Preference` variables, so it has no false positives on 10 000 lines. Testable via `-SourceRoots`, with self-tests in both directions |
| `tools/Invoke-PS51CompatibilityScan.ps1` | AST/token scan for PowerShell 7+-only constructs in code that must run on the declared **5.1 floor**: `??`, `??=`, ternary, `?.`, `clean{}`, `ConvertFrom-Json -AsHashtable`, `Join-Path -AdditionalChildPath`, and — the silent ones — `$IsWindows`/`$IsLinux` (which are `$null` on 5.1, so a branch fails *quietly*) and unguarded `ForEach-Object -Parallel` (which parses on 5.1 then fails at bind time). Token-based, so comments, help text, hashtable keys and the safe `Get-Variable -Name 'IsWindows'` probe are not findings |
| `tools/ps51-scanner-selftest.ps1` | 16-fixture self-test for the above: 10 constructs that **must** be flagged, 6 legitimate 5.1 patterns that **must not** — including the guarded `-Parallel` form the product genuinely relies on |
| `tools/Invoke-WorkflowLint.ps1` | Extracts every `run:` block from `.github/workflows/*.yml` and parse-validates it with the real PowerShell parser, plus the automatic-variable and failure-masking `exit 0` checks. YAML is opaque to `Invoke-Validation.ps1`, so without this a syntax error in CI would surface only on a GitHub runner — after the push |
| `tools/Invoke-DocLint.ps1` | Structural validation of all 34 Markdown documents: balanced code fences, table delimiter/column agreement, heading spacing, duplicate-heading anchor collisions, link hygiene, merge markers. Markdown has no compiler, so a damaged table renders as garbage on GitHub and in the wiki while every functional test stays green. Found two real defects — unescaped `\|` inside table cells, which GFM splits even inside backticks |
| `tools/Invoke-TemplateLint.ps1` | Validates `.github/**/*.yml` without a YAML-parser dependency (the project ships native-utilities-only). Catches invalid escapes in double-quoted scalars — which make GitHub render *"There was an error loading this form"* instead of the content — and issue-form schema violations (`markdown` needs `attributes.value`, every other type needs `attributes.label`). Block-scalar aware, so the PowerShell regexes embedded in `ci.yml` are not false positives |
| `tools/template-lint-selftest.ps1` | 12-fixture self-test for the above: 8 must be flagged, 4 legitimate patterns must not. This is what exposed defect 41 — the linter initially could not fail |
| `tools/Publish-Wiki.ps1` | Publishes `docs/wiki/` to `<repo>.wiki.git`, rewriting repository-file links into absolute GitHub URLs and validating all 25 pages' internal links first. Dry-run verified; the push needs your git credentials |

### The structural lesson

Every one of these defects was found by **adding a new observation point**, not by re-reading code:

- a second shipping form (module) needed its own gate stage → defects 21, 22, 23
- a release-integrity stage caught what functional tests could not see → defects 25, 26, 30, 34
- a hostile execution mode (StrictMode) caught what lax-mode tests could not → defect 29
- a static scanner for a language trap caught what no functional test could → defects 33, 35
- running the gate **nested** inside the build caught what a direct run could not → defects 27, 28
- a requirements-traceability stage caught what feature tests could not → defect 36
- a documentation linter caught what no code test can see → defects 37, 38
- reading the launcher's argument handling **semantically** (does every value-taking switch guard
  its value?) caught what 660 passing checks could not → defect 39, the only one of the forty-five
  that could have damaged a production fleet
- validating the repository's own GitHub metadata, which no product test ever reads → defect 40
- comparing a **normative document** to the bytes the code actually emits → defects 42, 43
- auditing an entire advertised surface rather than the one variable that happened to fail → defect 44
- running the same suite **both ways and diffing the output** → defect 45
- invoking the artifact **exactly as the README quotes it** - byte-for-byte, in a child process → defects 46, 47, 48

The gate must therefore be run both ways (`./tests/Run-Tests.ps1` and `./tools/Build.ps1`) before a
release; both are green - most recently from a clean clone of the published `v4` branch (Debian 12,
PowerShell 7.6.6, 2026-10-10).

---

## 6c. Release-verification pass (third sweep)

After the work tree was pushed to the public `v4` branch, the whole delivery was re-verified from a
**clean clone** on the platform CI uses for its Linux leg (Debian 12, PowerShell 7.6.6): `dist/`
deleted first (the defect-34 scenario), the gate run both ways, and then the artifact invoked
**exactly as the documentation quotes it**. That last step - the user-path review that caught
defect 39 in the launcher - found three more product defects, all in the help surface, all invisible
to 697 green checks:

* **46** - bare `-Help`, the README quick-start's own command, failed to bind, because `-Help` was a
  `[string]` section selector. Fixed: `-Help` is a switch rendering the full reference; the section
  selector moved to `-HelpSection <name>`. Comment-based help, the built-in reference, the launcher's
  forwarding, the README and five wiki/engineering pages were updated in the same change, so no
  surface still quotes the broken form.
* **47** - the standalone entry point's binding-error catch set `LASTEXITCODE = 3` and then let the
  next line overwrite it with RuntimeState's untouched "Success": **every malformed invocation exited
  0**, reporting success to schedulers and RMMs. Fixed: the catch owns the final code (3 =
  BadArguments, per the exit-code contract).
* **48** - under `Set-StrictMode -Version Latest` the full reference rendered as **silent empty
  output**: the Examples section read its optional note key as a property (`$e.N`), 5 of 15 entries
  lack the key, the strict-mode throw was swallowed by the product's own handler (exit 2, message
  only in the log). This is the defect-29 class surviving in the one document no strict test had
  rendered end-to-end. Fixed: indexer access for optional hashtable keys.

New permanent observation points: `tests/Unit.Help.ps1` (11 cases, auto-discovered by stage 5, runs
under the gate's strict pin - it renders the FULL document, which is what catches the 48 class) and
six new Verify-Brief checks that run the documented command lines **in a child `pwsh -File` process**
against the shipped artifact, including "invalid argument exits 3, never 0". Gate total grew
**697 -> 715 checks**, 9/9 stages, green on both paths, from a clean checkout.

The structural lesson, in one more form: *a command line quoted in documentation is a test
specification - run it byte-for-byte, in a child process, against the shipped artifact.*

---

## 7. To publish

State at handover (2026-10-10): the full v4 work tree is committed and pushed on the **`v4`
branch** of `paulmann/Print-Spooler-Queue-Cleanup` (the 83-file import commit, plus the
release-verification fixes of §6c on top). `main` still holds v3.1.0. **No `v4.0.0` tag exists
yet**, and `dist/` is git-ignored by design - the artifacts were built and verified locally
(Debian 12, PowerShell 7.6.6): gate 9/9, 715 checks, both paths green from a clean checkout.
Pushing to GitHub, tagging and creating the release need your credentials.

```powershell
# 1. Bring the release into main (GitHub PR from v4, or locally):
git checkout main
git merge --ff-only v4

# 2. Rebuild and re-verify (both paths must be green - see the note below)
./tools/Build.ps1 -Version 4.0.0        # add -SigningThumbprint <tp> to sign
./tests/Run-Tests.ps1                   # must be 9/9, 715 checks

# 3. Tag the release commit and push
git tag -a v4.0.0 -m "Print Spooler Queue Cleanup v4.0.0"
git push origin main --tags

# 4. Publish the wiki (dry-run first; it validates all 25 pages' links)
./tools/Publish-Wiki.ps1 -WhatIf
./tools/Publish-Wiki.ps1
```

**Run the gate both ways.** `./tests/Run-Tests.ps1` directly *and* via `./tools/Build.ps1` (which
nests the gate inside a build). Three defects in this project reproduced only on the nested path,
because that is the only one where a parent process had already exported environment variables and
left a `$LASTEXITCODE` behind.

Create the GitHub release from the `v4.0.0` tag and upload **everything** in `dist/`, including
`checksums.txt` — a client refuses to install an update whose checksums are missing. Paste
`dist/RELEASE-NOTES.md` as the release body.

Note that `dist/` is git-ignored on purpose: it is generated from `src/` on every build and must never
be hand-edited or committed. The release assets come from a local build, not from the repository.

The wiki is a **separate** git repository (`<repo>.wiki.git`) with a flat namespace, so
`tools/Publish-Wiki.ps1` handles the clone-modify-push cycle and rewrites repository-file links into
absolute GitHub URLs. GitHub only creates the wiki repository after you create the first page through
the web UI — create any page once, then let the script overwrite it.

If you would rather publish the wiki by hand, copy `docs/wiki/*.md` into the wiki editor; the pages
are already written for its flat namespace.

`.github/workflows/ci.yml` defines four jobs: the nine-stage gate on `windows-latest` **and**
`ubuntu-latest`; a Windows PowerShell **5.1** compatibility job; a PSScriptAnalyzer job; and a
`release` job that is `workflow_dispatch` only, because publishing a release is how every installed
client discovers the update and must never be a side effect of a merge.

The 5.1 job validates **both shipping forms** — it dot-sources the standalone script *and* separately
imports the module, re-running the `-SelfTest -PassThru` regression. That second step is what keeps
defect 21 from ever shipping again: the collision broke `-PassThru` only under `Import-Module`, so a
5.1 job that tested just the dot-source form would have stayed green through the exact regression it
exists to catch.

---

## 8. Known limitations (stated in the product, not hidden)

1. **Restoring spool files does not re-queue jobs.** `spoolsv.exe` builds its job table in memory.
   A restore is forensic recovery of the document payload for re-submission. Stated in
   `Restore-SpoolerBackup`'s output, the manifest's `RestoreNote`, the HTML report and the wiki.
2. **`.SHD` metadata is best-effort.** The header layout changed across Windows versions. When WMI is
   unavailable, name-based filters do **not** match jobs with unreadable metadata, and the run says so
   rather than guessing.
3. **`ForceStopFallback` terminates `spoolsv.exe`** — only after the normal stop has already timed
   out. In-flight jobs on that host are lost.
4. **The batch launcher cannot update itself while running** (`cmd.exe` re-reads by byte offset). Its
   own staged update lands on the next start, and it reports that.
5. **No reduced-privilege mode.** Administrator is required locally; a cleanup that cannot stop the
   Spooler would leave the queue worse than it found it.

---

## 9. Suggested next steps (not done, deliberately)

| Item | Status |
|---|---|
| Real-Windows integration test with an actual wedged queue | **Cannot be exercised here** (Linux sandbox, no Print Spooler). The payload is parse-validated, unit-tested and self-tested, but the first run on a live print server should go `-WhatIf` → `-Diagnostic` → `-Force -Backup` |
| Publishing the wiki to GitHub | `tools/Publish-Wiki.ps1` is ready and dry-run tested (25 pages, all links resolve). It needs your git credentials, so the push was not executed |
| Authenticode signing | Needs your code-signing certificate. `tools/Build.ps1 -SigningThumbprint <tp>` is wired and tested for absence of a cert |
| PSGallery publication | Needs a decision on the module name and a publishing account |
| Localization | Deliberately English-only per the brief |
| `PrintWorkflowUserSvc` per-user session handling over WinRM | Per-user services in other sessions need an interactive token; out of scope for a SYSTEM-run task |

# For Developers

How the project is built, tested and released — and the specific PowerShell traps that already bit
it, recorded so nobody has to rediscover them.

---

## Repository layout

```
src/                                  ← SOURCE OF TRUTH. Edit here only.
├── Shared/
│   ├── 00-Constants.ps1              metadata, services, registry paths, guards, exit codes, glyphs
│   └── 10-RuntimeState.ps1           the single mutable state object + initialiser
├── Private/
│   ├── 01-Platform.ps1               OS, elevation, console capability, spool-dir resolution
│   ├── 02-Logging.ps1                target resolution, sink, rotation, retention, Event Log
│   ├── 03-Console.ps1                rendering, formatting, confirmation gate, palette
│   ├── 04-Utility.ps1                versions, time spans, sizes, hashes, result factory, validation
│   ├── 05-Configuration.ps1          defaults, AST-safe file loading, merge, schema validation
│   ├── 06-RemotePayload.ps1          THE destructive logic, as a token-substituted template
│   ├── 07-Orchestration.ps1          token map, payload build, dispatch (local/remote/parallel)
│   ├── 08-BackupRestore.ps1          archive creation, manifest, restore, retention
│   ├── 09-Update.ps1                 discovery, checksums, staging, cooldown state
│   ├── 10-Report.ps1                 aggregation, console summary, JSON/CSV/HTML/Text export
│   ├── 11-Security.ps1               elevation, Authenticode, integrity, diagnostics bundle
│   ├── 12-ScheduledTask.ps1          task XML, registration via module or schtasks
│   ├── 13-HelpText.ps1               the built-in reference (19 sections)
│   └── 14-Execution.ps1              exit-code publishing, watchdog, self-test, arg reconstruction
└── Public/
    ├── Public-Cmdlets.ps1            the 14 exported commands
    └── Main-Function.ps1             Clear-PrintSpoolerQueue: param block + 9-step orchestration

bat/Clear-PrintSpoolerQueue.bat       launcher template (version token substituted at build)
config/PrintSpoolerCleanup.config.psd1   fully commented configuration template
tools/Build.ps1                       assembles dist/ from src/
tools/Invoke-Validation.ps1           parse gate
tools/Invoke-CollisionScan.ps1        parameter/local collision detector (self-tested)
tools/Invoke-AutomaticVariableScan.ps1 automatic-variable write detector (self-tested)
tools/Invoke-PS51CompatibilityScan.ps1 PS7-only syntax detector for the 5.1 floor (self-tested)
tools/module-test.ps1                 validates the .psd1/.psm1 form
tools/smoke-test.ps1                  end-to-end checks against the built artifact
tools/Verify-Release.ps1              audits the shippable artifacts (stage 8)
tools/Verify-Brief.ps1                proves every brief requirement still works (stage 9)
tools/Invoke-WorkflowLint.ps1         parse-validates the PowerShell embedded in CI YAML
tools/Invoke-TemplateLint.ps1         validates .github YAML (escapes, issue-form schema)
tools/Publish-Wiki.ps1                publishes docs/wiki/ to the GitHub wiki repository
tests/TestHarness.ps1                 dependency-free assert framework
tests/Unit.*.ps1                      unit suites
tests/Run-Tests.ps1                   the nine-stage quality gate (715 checks)
docs/                                 contracts and engineering notes
docs/wiki/                            the English wiki (mirrors the GitHub wiki)
.github/workflows/ci.yml              CI on windows-latest and ubuntu-latest
dist/                                 GENERATED. Never edit; never commit
```

---

## The build

```powershell
./tools/Build.ps1                       # build + validate + package + checksums
./tools/Build.ps1 -SkipTests -NoZip     # fast iteration
./tools/Build.ps1 -Version 4.1.0 -SigningThumbprint <thumbprint>
```

The build exists because v3 shipped **two hand-maintained copies** of the same logic — a standalone
script and the batch reimplementation — and they drifted, along with the README ("v3.0 Professional"
while the scripts said 3.1.0). v4 assembles *both* distributables from the same sources, so they
cannot disagree.

Steps:

1. Read `$Script:ScriptVersion` from `src/Shared/00-Constants.ps1` — the single source of truth
2. Concatenate `Shared → Private → Public` in a fixed order, wrapped in `#region`s
3. Substitute build-time tokens (`@@BUILD_STAMP@@`, `@@SCRIPT_VERSION@@`)
4. Emit `dist/Clear-PrintSpoolerQueue.ps1` = header + assembled + entry-point epilogue
5. Emit `dist/PrintSpoolerCleanup/PrintSpoolerCleanup.psm1` + a deterministic hand-written `.psd1`
   (not `New-ModuleManifest`, whose output varies by platform)
6. Stamp and encode the launcher: CRLF + pure ASCII
7. Copy the config template, README, LICENSE, CHANGELOG
8. **Validate**: parse every generated file; audit token substitution; check line endings and
   encoding; load the manifest and verify the export list
9. Generate `RELEASE-NOTES.md` from `CHANGELOG.md`
10. Package the ZIP and compute `checksums.txt`

### Two kinds of token — and why it matters

Build-time tokens (`@@BUILD_STAMP@@`) are substituted by the build. **Payload tokens**
(`@@SPOOLER_SERVICE@@`, `@@FORBIDDEN_ROOTS@@`, …) are substituted at *runtime* by
`Resolve-PayloadScriptBlock`, because the payload must arrive on a remote target as literal text.

The build therefore asserts a **whitelist** for the `.ps1`/`.psm1` (exactly the payload tokens may
survive) and **zero** tokens for the `.bat` and `.psd1`. An unknown surviving token fails the build,
which is a stronger check than "nothing left" and catches typos in new token names.

Substitution uses `[string]::Replace` (ordinal), never `-replace`: the latter is a regex operator and
both the pattern and the replacement would need escaping, since `$` and `\` are special in the
replacement string. A subtle `-replace` bug here would silently corrupt the payload on any site whose
paths contain a `$` — such as a hidden share `\\srv\Print$`.

---

## The payload architecture

The single most important design decision in the codebase.

```
New-PayloadOptions  →  flat, CLIXML-safe hashtable of scalars
Resolve-PayloadScriptBlock  →  token-substituted, parse-validated [scriptblock]
Invoke-CleanupFleet  →  local: direct invoke   |   remote: Invoke-Command
                         PS7: ForEach-Object -Parallel   |   PS5.1: runspace pool
ConvertFrom-PayloadResult  →  normalise into the public result schema
```

**One implementation runs everywhere.** v3 kept two divergent paths — local cmdlets versus
`Invoke-Command` — which is exactly how the v3.0 bug survived review where the remote scriptblock
referenced an undefined `$ScriptConfig.SpoolDirectory`. In v4, local and remote parity is structural
rather than aspirational.

Constraints on the payload, all deliberate:

- Only cmdlets available in **Windows PowerShell 5.1**, so the same code runs against Server 2012 R2
  and against PowerShell 7.6
- Only primitives, strings and string arrays in `$Options`, so everything survives CLIXML
  serialization
- Exactly one object on the output stream, normalised through a single `Finish-Result` exit point
  (generic `List[T]` does not always deserialize as expected; arrays do)
- No dependency on anything defined outside itself

---

## The quality gate

```powershell
./tests/Run-Tests.ps1              # everything
./tests/Run-Tests.ps1 -UnitOnly    # fast loop
./tests/Run-Tests.ps1 -SkipNetwork # offline / CI without egress
```

Nine stages, exit 0 only if all pass — **715 checks**:

| Stage | What | Checks |
|---|---|---|
| 1 | Build | `tools/Build.ps1` assembles both distributables |
| 2 | Parse validation | Every source **and** every generated artifact parses with zero errors |
| 3 | Source lint | 23 rules, each one a trap this codebase already hit |
| 4 | Batch lint | 14 rules: labels, gotos, shift counts, duplicate switches, ASCII, CRLF, exit codes, no own delete loop |
| 5 | Unit tests | **240** assertions against the built artifact, incl. the StrictMode suite |
| 6 | Module import | **36** assertions on the `.psd1`/`.psm1` form |
| 7 | Smoke test | **215** end-to-end assertions, in a separate process |
| 8 | Release integrity | **46** assertions: checksums, ZIP path-traversal safety, version coherence across all five surfaces, docs set, licence coherence |
| 9 | Brief compliance | **123** assertions that every requirement from the original brief is still present *and functional* |

Run the gate **both ways** — `./tests/Run-Tests.ps1` directly, and via `./tools/Build.ps1` (which
nests the gate inside a build). Three defects in this project reproduced only on the nested path,
because that is the only one where a parent process had already exported environment variables and
left a `$LASTEXITCODE` behind.

Unit tests run against `dist/`, not `src/`, so a build defect cannot hide behind a source-level pass.

**Stage 8 exists because "the code is correct" is not "the release is shippable".** It verifies that
every artifact's SHA-256 matches `checksums.txt`, that the ZIP contains what it should and nothing it
should not, that no ZIP entry can traverse outside the destination, that all five version stamps
agree, and that the README's licence claim matches `LICENSE` (finding A15 from the v3 audit).

**Stage 6 exists because of a real escape.** The module form is a genuinely separate code path —
module scope, `Export-ModuleMember`, private-helper isolation — and a `[switch]`-parameter collision
broke `-PassThru` *only* there. Four suites were green while the module was broken for everyone who
imported it. Any product that ships in two forms needs a gate for each form.

### No Pester

The product's promise is "native Windows utilities only, no external modules", and the test suite
holds itself to the same rule: it must run on a clean Windows machine with nothing installed and on a
Linux CI runner with only PowerShell. `tests/TestHarness.ps1` is ~150 lines of `Assert-*` functions
plus a runner. At this scale Pester would add a dependency and remove nothing.

### The source-lint rules

Each one exists because the bug actually happened:

| Rule | The bug it prevents |
|---|---|
| No `return switch` | `switch` is a statement, not an expression — a **syntax error** |
| No `.ContainsKey()` on ordered dictionaries | `[ordered]@{}` is an `OrderedDictionary`; only `.Contains()` exists. Every logging call threw |
| No nested `function` declarations | A nested function leaks its name into script scope once the build inlines the file |
| No mutation of a collection while enumerating its keys | See the cooldown bug below |
| No empty `catch { }` | An unexplained swallow is how the cooldown bug stayed invisible |
| No local variable colliding with a typed parameter | See trap 4(c) — broke `-PassThru` in the module form only |
| No unguarded `Measure-Object …).Sum` | See trap 16 — threw on every *empty* result set under StrictMode |
| No assignment to / splatting of an automatic variable | See trap 18 — `$args` shipped in both the CI release job and `-SelfElevate` |
| The gate pins its own `Set-StrictMode` | See trap 20 — an inherited strict mode made the verdict depend on the invocation path |
| No tabs | Batch tokenisation |
| Value-returning helpers must not `Write-Host` | A stray `Write-Host` corrupts the caller's pipeline |
| Every file parses | The substitute for a lexical quote-counting heuristic that produced 25 false positives |

---

## PowerShell traps that already bit this project

Recorded here because each one cost real debugging time and each one is invisible in review.

### 1. `return switch` is a syntax error

`switch` is a *statement*. `return switch ($x) { … }` fails to parse. Capture and return:

```powershell
$result = switch ($x) { 'a' { 1 } default { 2 } }
return $result
```

### 2. `$'` is a special variable

Inside a single-quoted string, a `$` immediately before the closing quote forms `$'` (text after the
last match) and **silently swallows the terminator**. The parse error then appears hundreds of lines
later, pointing at innocent code.

```powershell
'^m$|^min'        # BROKEN - the string does not terminate where you think
'^(?:m|min)$'     # also risky if the $ abuts the quote
'^(?:m|min)'      # fine - anchor at the start only
```

### 3. `switch` does not stop at the first match

…and `-regex` is a *substring* match. Unanchored, `'ms'` satisfied both the millisecond and the minute
clauses, and the function returned an **array of two TimeSpans** into a `[TimeSpan]` return value.

```powershell
$spans = @(switch -regex ($unit) {
    '^(?:ms|millisecond|milliseconds)$' { [TimeSpan]::FromMilliseconds($num) }
    '^(?:m|min|minute|minutes)$'        { [TimeSpan]::FromMinutes($num) }
})
if ($spans.Count -eq 0) { $spans = @([TimeSpan]::FromSeconds($num)) }
return [TimeSpan]$spans[0]
```

### 4. PowerShell variables are case-insensitive — including against *parameters*

The most dangerous trap in the language, and it bit this project **three times**.

**(a) Closures vs loop variables.** Closures named `$W`, `$P`, `$Sub` collided with loop variables
`$w`, `$p`, `$section` in the same scope and were **silently overwritten mid-render**. The symptom was
`The term 'W' is not recognized`, reported hundreds of lines away from the mistake. Use
multi-character names for anything that must survive a loop.

**(b) A loop variable vs a `[ValidateSet()]` parameter.** `$section` collided with the function's own
`$Section` parameter, so assigning the configuration section name `'Cleanup'` to it threw a
*validation* error about a value that was never user input.

**(c) A local variable vs a `[switch]` parameter.** In the entry point:

```powershell
param([switch]$SelfTest)
...
$selfTest = Invoke-SelfTestInternal     # <- this IS the $SelfTest parameter
```

Assigning a `PSCustomObject` to a `[switch]`-typed parameter throws
*"Cannot convert value … to type `System.Management.Automation.SwitchParameter`"*. Two properties made
this expensive to find:

- the throw happens **at the assignment**, so the message names an object that looks unrelated to the
  bug;
- it reproduced **only when the tool was loaded as a module**, not when dot-sourced — so 400+ checks
  across four suites stayed green while `-SelfTest -PassThru` was broken for every module user.

**Rule: never name a local variable after a parameter, in any casing.**
`tools/Invoke-CollisionScan.ps1` enforces this on every build, and the lint suite contains a
*self-test* that feeds the scanner a synthetic `[switch]` collision and asserts it fires — a detector
that cannot fail is not a detector.

### 5. `[ordered]@{}` has no `.ContainsKey()`

It is an `OrderedDictionary`, which implements `.Contains()`. Plain `@{}` has both. Using
`.Contains()` everywhere works for both — except `$PSBoundParameters`, which is a special dictionary
that has **only** `.ContainsKey()`.

### 6. `Mandatory` rejects empty strings before your guard runs

`[Parameter(Mandatory)][AllowNull()][string]$X` still refuses `''` at bind time. If an empty string is
a legitimate input — and for "no checksum was published" it is — you need `[AllowEmptyString()]`. Same
for `[AllowEmptyCollection()]` on arrays, where "every target was unreachable" is a real outcome.

### 7. A `[TimeSpan]`-typed parameter cannot receive `$null`

So `if ($null -eq $Duration)` inside the function is unreachable dead code. Leave the parameter
untyped and cast inside.

### 8. Enumerating `.Keys` while assigning into the same hashtable

```powershell
foreach ($k in $state.Keys) { $state[$k] = $json.$k }   # "Collection was modified"
foreach ($k in @($state.Keys)) { $state[$k] = $json.$k }   # snapshot first
```

Combined with an empty `catch {}`, this silently returned an **empty state**, which disabled the
update cooldown entirely and made every run hit the GitHub API. It is the most expensive one-line bug
in this project's history, and it is now both fixed and lint-guarded.

### 9. `GetFullPath()` is platform-relative

On Linux, `[IO.Path]::GetFullPath('C:\PRINTERS')` returns `<cwd>/C:\PRINTERS`. A depth guard built on
that result silently passed for a path that should have been rejected. Validate the drive-letter
shape **before** calling `GetFullPath`, and count segments from the input rather than the output.

### 10. `GetInvalidFileNameChars()` reflects the host

On Linux it omits `\`, `:` and `"`. A sanitiser built on it produces names that are invalid on the
Windows machines the tool actually targets. Carry the Windows set explicitly.

### 11. `String.Split(',', ';')` is not "split on both"

It binds to `Split(char separator, int count)` and **truncates**. Use `-split '[,;]'`.

### 12. `$$:` inside a double-quoted string

`"-$key`:`$$([string]$value)"` parses `$$` as the automatic variable (last token) followed by a colon,
emitting `-Name:` with no value. Build it by concatenation instead.

### 13. An unassigned value-returning call pollutes the pipeline

`Resolve-PayloadScriptBlock -Force` without `$null =` emitted a scriptblock into the output stream,
which then bound to the **next** command's parameters — surfacing as
`Cannot convert value … to type SwitchParameter` in a completely unrelated function.

### 14. `ConvertFrom-Json` converts ISO-8601 strings to `[datetime]`

PowerShell 7 does this automatically. A test asserting the *wire format* must check the raw line, not
the parsed property.

### 15. `$global:LASTEXITCODE` does not set the process exit code

Assigning it from inside a script does not change what `powershell -File` returns. Only `exit N` does
— and `exit N` from a dot-sourced script closes the operator's shell. Publish conditionally, based on
invocation context captured **in script scope** (`$MyInvocation` inside a function describes the
function, not the script).

### 16. `Measure-Object` returns `$null` for an empty pipeline

```powershell
$sum = ($list | Measure-Object -Property Bytes -Sum).Sum   # THROWS when $list is empty
```

`Measure-Object` with no input returns **`$null`**, not a zeroed `GenericMeasureInfo`. Reading `.Sum`
off `$null` is harmless in lax mode and a terminating error under `Set-StrictMode -Version Latest`.

There is a second, related hazard in the same expression family:

```powershell
$list | Where-Object { $_.BackupPath }     # THROWS if any record lacks BackupPath
```

Reading a property that a particular object does not carry is also a StrictMode error. Result records
here are deliberately **heterogeneous** — a `Cleanup` result has `JobsRemoved`, a `Diagnostic` result
does not — so aggregation must never assume a uniform shape.

Both are handled by one idiom:

```powershell
$prop = $Item.PSObject.Properties[$Name]     # $null when absent; never throws
if ($null -eq $prop) { return $null }
return $prop.Value
```

`Get-ResultPropertyValue` wraps that, and every sum and tally in the reporting layer goes through it.
A source-lint rule rejects any `Measure-Object …).Sum` in `src/`, and `tests/Unit.StrictMode.ps1`
re-runs aggregation and serialisation under `Set-StrictMode -Version Latest` with empty and
mixed-shape inputs.

Why this mattered: **an empty spool directory is the normal, healthy case**, not an edge case. Nine
sites summed with `Measure-Object` and six read properties directly, so a hardened host would have
failed on the most common possible run.

### 17. Tests must restore environment variables, not delete them

```powershell
$env:PSQ_LOG_PATH = 'X'
try { ... } finally { Remove-Item Env:\PSQ_LOG_PATH }   # WRONG
```

The quality gate exports `PSQ_LOG_PATH` so that no stage writes logs into the release artifact tree. A
test that *deletes* it un-redirects logging for **every later stage**, which then surfaces as three
failures in assertions that have nothing to do with logging. Use the harness helper
`Invoke-WithEnvironment`, which captures the prior state (including "did not exist") and restores it
exactly.

The same reasoning applies to `$LASTEXITCODE`: it is global mutable state that any earlier external
command can leave behind. Lint assertions that check a script's exit status now run it in a **child
process** and read `Process.ExitCode`.

Both problems were invisible when the gate was run directly and reproduced only when it ran nested
inside `Build.ps1` — which is why the gate must be exercised both ways before a release.

### 18. `$args` and friends are automatic variables, not free names

PowerShell maintains a set of variables itself. Assigning one is legal syntax and
usually *appears* to work:

```powershell
$args = @('release', 'create', $tag)     # legal... and wrong
gh @args
```

`$args` holds the **unbound arguments of the current invocation**. The host owns
it. Depending on scope and invocation style the assignment is overwritten, or it
silently consumes arguments the caller meant for you. The same applies to
`$Matches` (destroyed by the next `-match`), `$Error`, `$_` / `$PSItem`,
`$PSBoundParameters`, `$MyInvocation`, `$Switch` and `$foreach`.

This project shipped the bug **twice**: in the CI release job, and in
`New-ChildArgumentList` — the helper that builds the command line for
`-SelfElevate`. The latter worked for months only because it is always called with
*named* parameters; a single positional call would have returned the caller's
arguments instead of the child's.

`tools/Invoke-AutomaticVariableScan.ps1` fails the build on any assignment to or
splatting of these names. Two things make it usable rather than noise:

- it **excludes** `$null = ...` (the idiomatic output discard, used ~90 times here)
  and the `*Preference` variables (which exist precisely to be assigned);
- it accepts `-SourceRoots`, so the lint suite can point it at a fixture
  containing a known defect and assert it fires — and at a fixture of *correct*
  idioms and assert it stays quiet.

A detector with false positives gets disabled within a week, taking the real
protection with it. Both directions are tested.

### 19. The PowerShell 5.1 floor fails *silently* for two constructs

The product must run on Windows PowerShell 5.1 (Server 2012 R2 and Windows 7/8
fleets are still real). Most PS7-only syntax is a hard **parse** error there, which
is the good case: it fails immediately and loudly.

Two constructs do not:

```powershell
if ($IsWindows) { ... }        # $IsWindows is $null on 5.1 -> takes the ELSE branch
$items | ForEach-Object -Parallel { ... }   # parses on 5.1, fails at BIND time
```

`$IsWindows`, `$IsLinux` and `$IsMacOS` were introduced in PowerShell 6. On 5.1
they are simply undefined, so they evaluate to `$null` and the condition is false.
Nothing throws. The code takes the wrong branch and reports the wrong platform —
which for a tool that guards destructive filesystem operations is the worst
possible failure mode. Use `Test-RunningOnWindows` instead, which probes
`Get-Variable -Name 'IsWindows'` and falls back to `[Environment]::OSVersion`.

`-Parallel` is legitimate here — fleet dispatch genuinely needs it — but only
behind `Test-PowerShellSevenFeatures`, with the runspace-pool fallback for 5.1:

```powershell
if ($effectiveThrottle -gt 1 -and (Test-PowerShellSevenFeatures)) {
    $out = $targets | ForEach-Object -ThrottleLimit $effectiveThrottle -Parallel { ... }
}
else {
    # PS 5.1 path: runspace pool, same payload, same result shape
}
```

`tools/Invoke-PS51CompatibilityScan.ps1` enforces this at build time over `src/`
*and* the built artifacts. It is token/AST based rather than line-regex based,
which is what keeps it usable: the word `-Parallel` inside help text, `IsWindows`
as a hashtable key, and the safe `Get-Variable -Name 'IsWindows'` probe are all
correctly **not** findings. A regex scanner reports all three as defects, gets
ignored, and then the real one slips through.

Its self-test (`tools/ps51-scanner-selftest.ps1`) asserts both directions: 10
constructs that must be flagged, 6 legitimate 5.1 patterns that must not be.

### 20. `Set-StrictMode` leaks through `&`, and so does a test verdict

```powershell
# Build.ps1
Set-StrictMode -Version 2.0
& ./tests/Run-Tests.ps1 -Root $Root     # SAME process - strict mode is inherited
```

`&` (call operator) does **not** start a new process. `Set-StrictMode`, `$ErrorActionPreference` and
every other scope-visible setting carry straight into the called script. The consequence here was that
the identical test suite was **green when run directly and red when run through the build**, on
identical code:

```
./tests/Run-Tests.ps1     ->  240 passed, 0 failed     (lax)
./tools/Build.ps1         ->  239 passed, 1 failed     (strict, inherited)
```

A verdict that depends on how you invoked it is not a verdict. The fix is for the gate to pin its own
language mode rather than inherit whatever the caller happened to set:

```powershell
Set-StrictMode -Version Latest    # first thing in tests/Run-Tests.ps1
```

`Latest`, not `2.0` to match the build: the strictest setting makes "works only in lax mode" a build
failure instead of a production surprise. This is the same class as trap 16 — StrictMode as a
*hostile execution mode* used deliberately to surface latent defects.

**Run the gate both ways and diff the output.** That comparison is what found this; neither run alone
showed anything wrong.

### 21. `& $Body` runs in a child scope, so assignments do not escape

```powershell
function Assert-NotThrows { param([scriptblock]$Body) try { & $Body | Out-Null } catch { throw } }

$doc = $null
Assert-NotThrows { $doc = New-Object System.Xml.XmlDocument; $doc.LoadXml($x) }
$doc            # STILL $null - the assignment happened in the helper's child scope
```

The test then fails on `$doc` being null and the message points at the XML, not at the scoping. Capture
the error instead of relying on the assignment:

```powershell
$err = $null
try { $doc.LoadXml($x) } catch { $err = $_.Exception.Message }
Assert-Null $err 'must be well-formed'
```

Related: in `tests/TestHarness.ps1` every assert takes its `$Because` message **positionally**, so
`Assert-Equal 0 $x 'why'` reads naturally. Two of them originally omitted `Position`, which produced
*"A positional parameter cannot be found"* — a message that looks like a product defect and is
actually an inconsistency in the test framework.

---

## Adding a parameter

1. Add it to the `param()` block in `src/Public/Main-Function.ps1` with a `[ValidateRange]` /
   `[ValidateSet]` where one applies
2. Thread it into `$arguments` and, if the target needs it, into `New-PayloadOptions`
   (`src/Private/07-Orchestration.ps1`)
3. Document it in `$descriptions` inside `Get-ParameterDocumentation`
   (`src/Private/13-HelpText.ps1`) and add it to the right `$groups` entry
4. Add a comment-based `.PARAMETER` block if it deserves one
5. Add conflict rules to `Test-ParameterCombination` (`src/Private/04-Utility.ps1`) if any apply
6. Add a batch switch in `bat/Clear-PrintSpoolerQueue.bat` if it should be reachable from cmd
7. Add a config key in `Get-DefaultConfiguration` and a range check in `Test-Configuration` if it is
   a site-wide default
8. Add unit tests in `tests/Unit.*.ps1`
9. Mention it in `CLI-Reference.md`, the README table if user-facing, and `CHANGELOG.md`

Steps 3 and 8 are **test-enforced**: `-SelfTest` and the unit suite both fail if a parameter exists
without a description, or if a description exists for a parameter that does not.

---

## Adding a payload token

1. Add `@@YOUR_TOKEN@@` to the template in `src/Private/06-RemotePayload.ps1`
2. Add the mapping in `New-PayloadTokenMap` (`src/Private/07-Orchestration.ps1`). For an **array**
   value use `ConvertTo-QuotedListLiteral` and write `@(@@YOUR_TOKEN@@)` in the template — *not*
   `@('@@YOUR_TOKEN@@')`, which produces `@(''a', 'b'')` and fails to parse
3. Add it to `$payloadTokenWhitelist` in `tools/Build.ps1`
4. Add an assertion in `tests/Unit.Guards.ps1` proving it substituted

Forgetting step 3 fails the build with "unexpected tokens", which is the intended outcome.

---

## Release checklist

- [ ] `src/Shared/00-Constants.ps1` → bump `$Script:ScriptVersion`
- [ ] `CHANGELOG.md` → add the version section (Added / Changed / Fixed / Removed)
- [ ] `README.md` → update the version badge and any changed examples
- [ ] `./tools/Build.ps1 -Version <v>` → confirm the gate passes
- [ ] `./tests/Run-Tests.ps1` → 9/9 stages green (run it again via `./tools/Build.ps1` — three
      defects in this project reproduced only on the nested path)
- [ ] If you touched anything that runs remotely or on old hosts, confirm the 5.1 floor held:
      `./tools/Invoke-PS51CompatibilityScan.ps1 -IncludeArtifacts` (stage 3 runs it automatically)
- [ ] Review `dist/RELEASE-NOTES.md` (generated from the CHANGELOG)
- [ ] Sign if you publish signed artifacts: `-SigningThumbprint <thumbprint>`
- [ ] Push a `v<version>` tag
- [ ] Create the GitHub release and upload **every** file in `dist/`, including `checksums.txt`
      (clients refuse to install an update whose checksums are missing)
- [ ] Update the GitHub wiki: `./tools/Publish-Wiki.ps1 -WhatIf` then without `-WhatIf`
      (the wiki is a *separate* git repository with a flat namespace; the tool handles the
      clone-modify-push cycle and rewrites repository-file links into absolute URLs)

The version stamp is read from one constant and written into the `.ps1`, `.psd1`, `.bat` and release
notes by the build, which fails if any `@@TOKEN@@` survives. That is the mechanism that prevents the
v3 situation where the README, the script and the launcher all claimed different versions.

---

## CI

`.github/workflows/ci.yml` defines four jobs. Every push and pull request runs the first three; the
fourth is manual.

| Job | Runs on | What it proves |
|---|---|---|
| `quality-gate` | `windows-latest` **and** `ubuntu-latest`, PowerShell 7 | The full nine-stage gate, 715 checks, on both platforms |
| `powershell-51` | `windows-latest`, Windows PowerShell **5.1** | The built artifact parses, dot-sources, imports and self-tests on the *oldest* supported engine |
| `static-analysis` | `ubuntu-latest` | PSScriptAnalyzer over the sources (informational, not blocking) |
| `release` | `windows-latest`, `workflow_dispatch` only | Builds, tests, packages and publishes a GitHub release |

### Why the Linux leg is not decoration

It caught the `GetFullPath('C:\PRINTERS')` platform-relative bug — a guard that silently passed on the
developer's Windows machine and would have been meaningless anywhere else. Platform-dependent
behaviour in a **safety guard** is exactly what a cross-platform matrix exists to find. It also caught
`Join-Path` throwing `Cannot find drive 'C'` on a host with no C: drive.

### Why the 5.1 job validates *both* shipping forms

The 5.1 job dot-sources the standalone script **and** separately imports the module. That second step
is not redundant, and the reason is defect 21 in the audit: a local `$selfTest` colliding with the
`[switch]$SelfTest` parameter broke `-PassThru` *only* when the tool was loaded via `Import-Module`.
Dot-sourcing, the parser, four test suites and 400+ checks were all green while every module user was
broken.

Module scope, `Export-ModuleMember` and private-helper isolation are a genuinely separate code path,
so each form gets its own gate — locally (`tools/module-test.ps1`, stage 6) and in CI, on both engines.

### Why `release` is manual

Publishing a release is how **every installed client discovers the update**, through the version-check
subsystem. A release must never be a side effect of merging a pull request, so the job is
`workflow_dispatch` only.

---

## Contributing

See [CONTRIBUTING.md](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/blob/main/CONTRIBUTING.md).
The short version:

1. Edit `src/`, never `dist/`
2. `./tests/Run-Tests.ps1` must pass
3. New behaviour needs a test; a fixed bug needs a test that failed before
4. English comments and documentation, explaining **why** rather than restating the code
5. When you hit a PowerShell trap, add it to the list above and, if it can recur, to the source lint

---

## Related pages

- [Safety Model](Safety-Model)
- [Updates and Versioning](Updates-and-Versioning)
- [CLI Reference](CLI-Reference)
- [Cmdlet Reference](Cmdlet-Reference)

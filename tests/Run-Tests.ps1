#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Full quality gate: build, parse-validate, unit tests, batch lint, smoke tests.

.DESCRIPTION
    Runs everything that must be true before a release is published:

      1. tools/Build.ps1 assembles the distributables from src/
      2. every source file AND every generated artifact parses with zero errors
      3. the unit suite runs against the BUILT artifact (not the sources), so a
         build defect cannot hide behind a source-level pass
      4. the batch launcher is linted structurally (labels, gotos, encoding,
         balanced parentheses, no tabs)
      5. tools/smoke-test.ps1 runs end-to-end against the built artifact
      6. the source lint catches the specific PowerShell traps this codebase has
         already been bitten by

    Exit code 0 only when every stage passes.

.PARAMETER SkipBuild
    Test the existing dist/ as-is.

.PARAMETER SkipNetwork
    Skip the live GitHub update-discovery checks.

.EXAMPLE
    ./tests/Run-Tests.ps1
    ./tests/Run-Tests.ps1 -SkipNetwork
#>
[CmdletBinding()]
param(
    [string]$Root,
    [switch]$SkipBuild,
    [switch]$SkipNetwork,
    [switch]$UnitOnly
)

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($Root)) { $Root = Split-Path -Parent $PSScriptRoot }
$Root = (Resolve-Path -LiteralPath $Root).Path

$stageFailures = 0
function Write-Stage { param([string]$T) Write-Host "`n$('=' * 78)`n  $T`n$('=' * 78)" -ForegroundColor Cyan }

# Pin the language mode EXPLICITLY. tools/Build.ps1 sets Set-StrictMode -Version
# 2.0 and then invokes this script with `&`, i.e. in the SAME process, so the
# strict mode used to leak into the whole suite. The result was that the identical
# suite passed when run directly (lax) and failed when run via Build.ps1 (strict) -
# a test outcome that depends on how you invoked it is not a test result.
#
# Latest is chosen deliberately rather than matching Build's 2.0: it is the
# strictest setting, so a test that only works in lax mode is caught here instead
# of in production. Defect 29 (the product crashing under StrictMode) is exactly
# the class this surfaces.
Set-StrictMode -Version Latest

# Redirect ALL logging away from the artifact tree for the whole gate.
#
# Stages 5-7 dot-source and import the built artifacts, which resolves a log
# target. Where %ProgramData% is unavailable (Linux CI, a locked-down host) the
# fallback chain reaches the script's own directory - i.e. dist/ - so exercising
# the product would write Logs/ INTO the release artifacts, after Build.ps1 had
# already swept them. Setting PSQ_LOG_PATH once here fixes the ordering at its
# root, is inherited by every child process the gate spawns, and incidentally
# proves that PSQ_LOG_PATH works end to end.
$savedLogEnv = $env:PSQ_LOG_PATH
$gateLogDir = Join-Path ([IO.Path]::GetTempPath()) ("psq-gate-{0}" -f [Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $gateLogDir -Force
$env:PSQ_LOG_PATH = Join-Path $gateLogDir 'quality-gate.log'

# ---------------------------------------------------------------------------
# Stage 1: build
# ---------------------------------------------------------------------------
if (-not $SkipBuild -and -not $UnitOnly) {
    Write-Stage 'STAGE 1/9  Build the distributables from src/'
    # NOTE: no -NoZip here. Stage 8 verifies RELEASE artifact integrity, which
    # includes the ZIP: its contents, CRC, path-traversal safety and whether it
    # is in sync with the sources. Building without the archive meant stage 8
    # either validated a STALE zip left behind by an earlier build, or - on a
    # clean checkout, which is exactly what CI starts from - threw an unhandled
    # MethodInvocationException trying to open an archive that did not exist.
    & (Join-Path $Root 'tools/Build.ps1') -Root $Root -SkipTests
    if ($LASTEXITCODE -ne 0) { Write-Host 'Build failed.' -ForegroundColor Red; exit 1 }
}

$dist = Join-Path $Root 'dist'
$artifact = Join-Path $dist 'Clear-PrintSpoolerQueue.ps1'
if (-not (Test-Path -LiteralPath $artifact)) { throw "No built artifact at $artifact - run without -SkipBuild first." }

# ---------------------------------------------------------------------------
# Stage 2: parse validation
# ---------------------------------------------------------------------------
if (-not $UnitOnly) {
    Write-Stage 'STAGE 2/9  Parse validation (sources + artifacts)'
    & (Join-Path $Root 'tools/Invoke-Validation.ps1') -Root $Root
    if ($LASTEXITCODE -ne 0) { $stageFailures++; Write-Host 'Parse validation FAILED.' -ForegroundColor Red }
}

# ---------------------------------------------------------------------------
# Stage 3: source lint for known PowerShell traps
# ---------------------------------------------------------------------------
Write-Stage 'STAGE 3/9  Source lint (PowerShell traps this codebase has hit)'
. (Join-Path $Root 'tests/TestHarness.ps1')
Reset-TestCounters
Start-TestSuite 'Source lint'

$sourceFiles = @(Get-ChildItem -LiteralPath (Join-Path $Root 'src') -Recurse -Filter *.ps1 -File | Sort-Object FullName)

Invoke-TestCase -Name 'No "return switch" (switch is a statement, not an expression)' -Body {
    $bad = @()
    foreach ($f in $sourceFiles) {
        $n = 0
        foreach ($line in (Get-Content -LiteralPath $f.FullName)) {
            $n++
            $code = ($line -replace '#.*$', '')          # comments may legitimately quote the trap
            if ($code -match '\breturn\s+switch\b') { $bad += "$($f.Name):$n" }
        }
    }
    Assert-Equal 0 $bad.Count ("offenders: " + ($bad -join ', '))
}

Invoke-TestCase -Name 'Every source file parses with zero errors' -Body {
    # This is the sound version of the trap-hunt: an unterminated string caused
    # by a '$' before a closing quote is a PARSE error, so the parser catches it
    # exactly, with no false positives. A lexical quote-counting heuristic was
    # tried first and flagged 25 legitimate lines (apostrophes inside
    # double-quoted text), so it was removed rather than tuned.
    $bad = @()
    foreach ($f in $sourceFiles) {
        $t = $null; $e = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$t, [ref]$e)
        if ($e -and $e.Count -gt 0) { $bad += "$($f.Name) ($($e.Count))" }
    }
    Assert-Equal 0 $bad.Count ("parse errors in: " + ($bad -join ', '))
}

Invoke-TestCase -Name 'No mutation of a collection while enumerating its keys' -Body {
    # `foreach ($k in $x.Keys) { $x[$k] = ... }` throws "Collection was modified"
    # part-way through. Inside a try/catch that reports nothing, it silently
    # returns half-built data - which is exactly how the update cooldown was
    # disabled without any visible symptom.
    $bad = @()
    $pattern = [regex]'foreach\s*\(\s*\$(\w+)\s+in\s+\$(\w+(?:\[[^\]]+\])*)\.Keys\s*\)'
    foreach ($f in $sourceFiles) {
        $lines = Get-Content -LiteralPath $f.FullName
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $m = $pattern.Match($lines[$i])
            if (-not $m.Success) { continue }
            $coll = [regex]::Escape($m.Groups[2].Value)
            for ($j = $i; $j -lt [Math]::Min($i + 9, $lines.Count); $j++) {
                if ($lines[$j] -match "\`$$coll\[[^\]]+\]\s*=") { $bad += "$($f.Name):$($i + 1)"; break }
            }
        }
    }
    Assert-Equal 0 $bad.Count ("mutation during enumeration at: " + ($bad -join ', '))
}

Invoke-TestCase -Name 'No empty catch blocks that can hide a failure' -Body {
    # `catch { }` with no comment is how the defect above stayed invisible. An
    # intentional swallow must say why, so require a comment inside the block.
    $bad = @()
    foreach ($f in $sourceFiles) {
        $lines = Get-Content -LiteralPath $f.FullName
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match 'catch\s*\{\s*\}') { $bad += "$($f.Name):$($i + 1)" }
        }
    }
    Assert-Equal 0 $bad.Count ("silent catch at: " + ($bad -join ', '))
}

Invoke-TestCase -Name 'No statement precedes param() at script scope' -Body {
    # In PowerShell a script's param() block - and any [CmdletBinding()] attribute
    # on it - must be the FIRST thing in the file. A single assignment before it
    # is a hard parse error ("Unexpected attribute 'CmdletBinding'"), and it is
    # remarkably easy to write: `$root = Split-Path -Parent $PSScriptRoot` above
    # the comment-based help reads naturally and breaks the file.
    #
    # This was written THREE times while adding tools to this project. It is
    # checked with the AST, not by scanning lines: a regex over the text matches
    # prose inside comment-based help blocks and reports ~10 false positives.
    $bad = @()
    foreach ($f in $sourceFiles) {
        $tokens = $null; $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$tokens, [ref]$errors)
        if (-not $ast -or -not $ast.ParamBlock) { continue }
        $pbStart = $ast.ParamBlock.Extent.StartOffset
        foreach ($st in $ast.EndBlock.Statements) {
            if ($st.Extent.StartOffset -lt $pbStart) {
                $bad += ("{0}:{1}" -f $f.Name, $st.Extent.StartLineNumber)
            }
        }
    }
    Assert-Equal 0 $bad.Count $(if ($bad) { 'statement before param(): ' + ($bad -join ', ') } else { 'all param blocks are first' })
}

Invoke-TestCase -Name 'Embedded PowerShell in CI workflows parses' -Body {
    # The workflow embeds ~140 lines of PowerShell inside YAML `run:` blocks.
    # tools/Invoke-Validation.ps1 parses every *.ps1, but YAML is opaque to it, so
    # a syntax error there is invisible locally and only surfaces on a GitHub
    # runner - after the push. This also re-checks the automatic-variable rule
    # inside YAML, where the .ps1 scanner cannot see.
    $lint = Join-Path $Root 'tools/Invoke-WorkflowLint.ps1'
    Assert-True (Test-Path -LiteralPath $lint) 'the workflow linter must exist'
    $exe = (Get-Process -Id $PID).Path
    $outFile = Join-Path ([IO.Path]::GetTempPath()) ("psq-wf-{0}.out" -f [Guid]::NewGuid().ToString('N'))
    $proc = Start-Process -FilePath $exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $lint) `
        -NoNewWindow -Wait -PassThru -RedirectStandardOutput $outFile
    $out = if (Test-Path -LiteralPath $outFile) { Get-Content -LiteralPath $outFile -Raw } else { '' }
    Remove-Item -LiteralPath $outFile -Force -ErrorAction SilentlyContinue
    Assert-Equal 0 $proc.ExitCode ('a CI workflow block failed to lint: ' + $out)
    Assert-Match 'WORKFLOW LINT' $out 'the linter must report its summary'
}

Invoke-TestCase -Name 'No automatic variable is assigned or splatted' -Body {
    # $args, $Matches, $Error, $_, $PSBoundParameters and friends are maintained
    # by the host. Assigning one appears to work, then the host repopulates it and
    # the value vanishes - or it silently corrupts the caller's state. Two
    # instances of the $args form shipped here before this rule existed: the CI
    # release job splatted `gh @args`, and the production New-ChildArgumentList
    # helper built its argument vector in `$args`.
    #
    # $null = ... (output discard) and the *Preference variables are correct and
    # pervasive, so the scanner excludes them by design.
    $scan = Join-Path $Root 'tools/Invoke-AutomaticVariableScan.ps1'
    Assert-True (Test-Path -LiteralPath $scan) 'the automatic-variable scanner must exist'
    # Child process, not $LASTEXITCODE: see the collision rule for why.
    $exe = (Get-Process -Id $PID).Path
    $outFile = Join-Path ([IO.Path]::GetTempPath()) ("psq-auto-{0}.out" -f [Guid]::NewGuid().ToString('N'))
    $proc = Start-Process -FilePath $exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scan) `
        -NoNewWindow -Wait -PassThru -RedirectStandardOutput $outFile
    $out = if (Test-Path -LiteralPath $outFile) { Get-Content -LiteralPath $outFile -Raw } else { '' }
    Remove-Item -LiteralPath $outFile -Force -ErrorAction SilentlyContinue
    Assert-Equal 0 $proc.ExitCode ('an automatic variable is assigned or splatted: ' + $out)
}

Invoke-TestCase -Name 'The automatic-variable scanner detects a synthetic $args collision' -Body {
    # A detector that cannot fail is not a detector. The earlier version of this
    # case re-implemented the scanner's AST walk inline, which proved nothing
    # about the scanner itself. This one runs the REAL scanner against a fixture
    # tree through its -SourceRoots parameter and asserts on its exit code and
    # output - the same contract the production rule above depends on.
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("psq-auto-bad-{0}" -f [Guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $tmp -Force
    try {
        $bad = @'
function Test-BadArgs {
    [CmdletBinding()]
    param([string[]]$Items)
    $args = New-Object System.Collections.Generic.List[string]
    foreach ($i in $Items) { $args.Add($i) }
    return $args.ToArray()
}
'@
        Set-Content -LiteralPath (Join-Path $tmp 'Bad.ps1') -Value $bad -Encoding UTF8

        $exe = (Get-Process -Id $PID).Path
        $scan = Join-Path $Root 'tools/Invoke-AutomaticVariableScan.ps1'
        $outFile = Join-Path ([IO.Path]::GetTempPath()) ("psq-auto-bad-{0}.out" -f [Guid]::NewGuid().ToString('N'))
        $proc = Start-Process -FilePath $exe `
            -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scan, '-SourceRoots', $tmp, '-SkipWorkflow') `
            -NoNewWindow -Wait -PassThru -RedirectStandardOutput $outFile
        $out = if (Test-Path -LiteralPath $outFile) { Get-Content -LiteralPath $outFile -Raw } else { '' }
        Remove-Item -LiteralPath $outFile -Force -ErrorAction SilentlyContinue

        Assert-Equal 1 $proc.ExitCode 'the scanner failed to flag an assignment to $args'
        Assert-Match 'AUTOMATIC VARIABLE' $out 'the scanner must name the finding class'
        Assert-Match 'args' $out 'the scanner must name the offending variable'
    }
    finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

Invoke-TestCase -Name 'The automatic-variable scanner ignores correct idioms' -Body {
    # The other half of a useful detector: it must NOT fire on `$null = ...`
    # (output discard) or on the *Preference variables, both of which are correct
    # and appear hundreds of times in this codebase. A rule that flags them would
    # be disabled within a week, taking the real protection with it.
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("psq-auto-ok-{0}" -f [Guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $tmp -Force
    try {
        $ok = @'
function Test-OkIdioms {
    [CmdletBinding()]
    param([string]$Path)
    $ErrorActionPreference = 'Stop'
    $ProgressPreference = 'SilentlyContinue'
    $null = New-Item -ItemType Directory -Path $Path -Force
    $items = @()
    foreach ($f in @(Get-ChildItem -LiteralPath $Path -File)) { $items += $f.Name }
    return $items
}
'@
        Set-Content -LiteralPath (Join-Path $tmp 'Ok.ps1') -Value $ok -Encoding UTF8

        $exe = (Get-Process -Id $PID).Path
        $outFile = Join-Path ([IO.Path]::GetTempPath()) ("psq-auto-ok-{0}.out" -f [Guid]::NewGuid().ToString('N'))
        $proc = Start-Process -FilePath $exe `
            -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $Root 'tools/Invoke-AutomaticVariableScan.ps1'), '-SourceRoots', $tmp, '-SkipWorkflow', '-Quiet') `
            -NoNewWindow -Wait -PassThru -RedirectStandardOutput $outFile
        $out = if (Test-Path -LiteralPath $outFile) { Get-Content -LiteralPath $outFile -Raw } else { '' }
        Remove-Item -LiteralPath $outFile -Force -ErrorAction SilentlyContinue

        Assert-Equal 0 $proc.ExitCode ('the scanner raised a false positive on correct idioms: ' + $out)
    }
    finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

Invoke-TestCase -Name 'Sources and built artifacts are PowerShell 5.1 compatible' -Body {
    # 5.1 is the declared floor for the product. Most PS7-only syntax is a hard
    # PARSE error there, which is loud. Two categories are not:
    #
    #   * $IsWindows / $IsLinux / $IsMacOS are simply $null on 5.1, so
    #     `if ($IsWindows)` takes the WRONG branch silently instead of failing.
    #   * ForEach-Object -Parallel parses on 5.1 and then fails at BIND time, on
    #     whichever fleet machine happens to be old. It is legitimate here, but
    #     only behind Test-PowerShellSevenFeatures with the runspace fallback.
    #
    # Detection is AST/token based, so comments, help text, hashtable keys
    # (`IsWindows = ...`) and the safe `Get-Variable -Name 'IsWindows'` probe are
    # correctly NOT findings - a regex scanner reports all four as defects.
    $scan = Join-Path $Root 'tools/Invoke-PS51CompatibilityScan.ps1'
    Assert-True (Test-Path -LiteralPath $scan) 'the 5.1 compatibility scanner must exist'
    $exe = (Get-Process -Id $PID).Path
    $outFile = Join-Path ([IO.Path]::GetTempPath()) ("psq-51-{0}.out" -f [Guid]::NewGuid().ToString('N'))
    $proc = Start-Process -FilePath $exe `
        -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scan, '-IncludeArtifacts') `
        -NoNewWindow -Wait -PassThru -RedirectStandardOutput $outFile
    $out = if (Test-Path -LiteralPath $outFile) { Get-Content -LiteralPath $outFile -Raw } else { '' }
    Remove-Item -LiteralPath $outFile -Force -ErrorAction SilentlyContinue
    Assert-Equal 0 $proc.ExitCode ('PowerShell 7+-only construct in shipped code: ' + $out)
}

Invoke-TestCase -Name 'The 5.1 compatibility scanner is self-tested in both directions' -Body {
    # A detector that cannot fail is not a detector, and one that only ever fires
    # is not useful either - it gets disabled, taking the real protection with it.
    # This runs the fixture suite: 10 constructs that MUST be flagged, and 6
    # legitimate 5.1 patterns that MUST NOT be, including the guarded -Parallel
    # form the product genuinely relies on.
    $selfTest = Join-Path $Root 'tools/ps51-scanner-selftest.ps1'
    Assert-True (Test-Path -LiteralPath $selfTest) 'the 5.1 scanner self-test must exist'
    $exe = (Get-Process -Id $PID).Path
    $outFile = Join-Path ([IO.Path]::GetTempPath()) ("psq-51st-{0}.out" -f [Guid]::NewGuid().ToString('N'))
    $proc = Start-Process -FilePath $exe `
        -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $selfTest) `
        -NoNewWindow -Wait -PassThru -RedirectStandardOutput $outFile
    $out = if (Test-Path -LiteralPath $outFile) { Get-Content -LiteralPath $outFile -Raw } else { '' }
    Remove-Item -LiteralPath $outFile -Force -ErrorAction SilentlyContinue
    Assert-Equal 0 $proc.ExitCode ('the 5.1 scanner self-test failed: ' + $out)
    Assert-Match 'FIRES' $out 'the self-test must prove the scanner fires on PS7-only code'
    Assert-Match 'QUIET' $out 'the self-test must prove the scanner stays quiet on 5.1 code'
}

Invoke-TestCase -Name 'Markdown documents are structurally valid' -Body {
    # These documents were edited programmatically many times: rows inserted into
    # tables, sections replaced, counters rewritten. Markdown has no compiler, so
    # a damaged table renders as garbage on GitHub and in the wiki and NOTHING
    # else in the gate would ever notice.
    #
    # Two real defects were found this way, both invisible to every other check:
    # an unescaped `|` inside a table cell (GFM splits cells on `|` even inside
    # backticks, so one row rendered with an extra column) and an unclosed code
    # fence would swallow the rest of a file.
    $lint = Join-Path $Root 'tools/Invoke-DocLint.ps1'
    Assert-True (Test-Path -LiteralPath $lint) 'the documentation linter must exist'
    $exe = (Get-Process -Id $PID).Path
    $outFile = Join-Path ([IO.Path]::GetTempPath()) ("psq-doc-{0}.out" -f [Guid]::NewGuid().ToString('N'))
    $proc = Start-Process -FilePath $exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $lint, '-Root', $Root, '-Quiet') `
        -NoNewWindow -Wait -PassThru -RedirectStandardOutput $outFile
    $out = if (Test-Path -LiteralPath $outFile) { Get-Content -LiteralPath $outFile -Raw } else { '' }
    Remove-Item -LiteralPath $outFile -Force -ErrorAction SilentlyContinue
    Assert-Equal 0 $proc.ExitCode ('a Markdown document is structurally broken: ' + $out)
}

Invoke-TestCase -Name 'The documentation linter detects broken tables and unbalanced fences' -Body {
    # Both directions: it must fire on damage, and it must NOT fire on the '='
    # banners inside console-output code blocks that this project's documentation
    # is full of. A linter that flags sample transcripts gets disabled.
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("psq-doc-bad-{0}" -f [Guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $tmp -Force
    try {
        $exe = (Get-Process -Id $PID).Path
        $lint = Join-Path $Root 'tools/Invoke-DocLint.ps1'

        # (a) a table row with the wrong column count
        Set-Content -LiteralPath (Join-Path $tmp 'Bad-Table.md') -Encoding UTF8 -Value @'
# Title

| A | B |
|---|---|
| 1 | 2 | 3 |
'@
        $o1 = Join-Path ([IO.Path]::GetTempPath()) ("psq-doc-o1-{0}.out" -f [Guid]::NewGuid().ToString('N'))
        $p1 = Start-Process -FilePath $exe -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',$lint,'-Root',$tmp,'-Quiet') -NoNewWindow -Wait -PassThru -RedirectStandardOutput $o1
        $t1 = if (Test-Path -LiteralPath $o1) { Get-Content -LiteralPath $o1 -Raw } else { '' }
        Remove-Item -LiteralPath $o1 -Force -ErrorAction SilentlyContinue
        Assert-Equal 1 $p1.ExitCode 'the linter failed to flag a malformed table row'

        # (b) an unbalanced code fence
        Remove-Item -LiteralPath (Join-Path $tmp 'Bad-Table.md') -Force
        Set-Content -LiteralPath (Join-Path $tmp 'Bad-Fence.md') -Encoding UTF8 -Value "# Title`n`n``````powershell`nGet-Date`n"
        $o2 = Join-Path ([IO.Path]::GetTempPath()) ("psq-doc-o2-{0}.out" -f [Guid]::NewGuid().ToString('N'))
        $p2 = Start-Process -FilePath $exe -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',$lint,'-Root',$tmp,'-Quiet') -NoNewWindow -Wait -PassThru -RedirectStandardOutput $o2
        Remove-Item -LiteralPath $o2 -Force -ErrorAction SilentlyContinue
        Assert-Equal 1 $p2.ExitCode 'the linter failed to flag an unclosed code fence'

        # (c) correct Markdown that CONTAINS '=' banners inside a fence must pass
        Remove-Item -LiteralPath (Join-Path $tmp 'Bad-Fence.md') -Force
        Set-Content -LiteralPath (Join-Path $tmp 'Good.md') -Encoding UTF8 -Value @'
# Title

```
==============================================================================
    EXECUTION SUMMARY
==============================================================================
```

| A | B |
|---|---|
| 1 | 2 |
| `x \| y` | escaped pipe is fine |
'@
        $o3 = Join-Path ([IO.Path]::GetTempPath()) ("psq-doc-o3-{0}.out" -f [Guid]::NewGuid().ToString('N'))
        $p3 = Start-Process -FilePath $exe -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',$lint,'-Root',$tmp,'-Quiet') -NoNewWindow -Wait -PassThru -RedirectStandardOutput $o3
        $t3 = if (Test-Path -LiteralPath $o3) { Get-Content -LiteralPath $o3 -Raw } else { '' }
        Remove-Item -LiteralPath $o3 -Force -ErrorAction SilentlyContinue
        Assert-Equal 0 $p3.ExitCode ('the linter raised a false positive on valid Markdown: ' + $t3)
    }
    finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

Invoke-TestCase -Name 'GitHub YAML templates are valid and render' -Body {
    # Two defect classes here fail SILENTLY and are invisible to every other check:
    #
    #   1. An invalid escape in a DOUBLE-QUOTED YAML scalar. `".\Clear-..."` is a
    #      hard parse error, and GitHub then renders "There was an error loading
    #      this form" instead of the content. Nothing fails at build time.
    #   2. An issue-form element that does not match GitHub's schema - a `markdown`
    #      element carries its text in `attributes.value`, every other type needs
    #      `attributes.label`. Wrong key => the element renders empty.
    #
    # The escape check must be block-scalar aware: inside `run: |` the content is
    # literal, so the PowerShell regexes embedded in ci.yml legitimately contain
    # backslashes that are NOT YAML escapes. A naive line scanner reports six false
    # positives there and gets switched off.
    $lint = Join-Path $Root 'tools/Invoke-TemplateLint.ps1'
    Assert-True (Test-Path -LiteralPath $lint) 'the template linter must exist'
    $exe = (Get-Process -Id $PID).Path
    $outFile = Join-Path ([IO.Path]::GetTempPath()) ("psq-tpl-{0}.out" -f [Guid]::NewGuid().ToString('N'))
    $proc = Start-Process -FilePath $exe `
        -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $lint, '-Root', $Root) `
        -NoNewWindow -Wait -PassThru -RedirectStandardOutput $outFile
    $out = if (Test-Path -LiteralPath $outFile) { Get-Content -LiteralPath $outFile -Raw } else { '' }
    Remove-Item -LiteralPath $outFile -Force -ErrorAction SilentlyContinue
    Assert-Equal 0 $proc.ExitCode ('a GitHub YAML template would not render: ' + $out)
}

Invoke-TestCase -Name 'The template linter is self-tested in both directions' -Body {
    # It initially could not fail at all: it resolved the tree from its own
    # location, so pointing it at a fixture via -WorkingDirectory was ignored and
    # every broken fixture "passed". The self-test is what exposed that. Eight
    # fixtures must be flagged, four legitimate patterns must not be.
    $selfTest = Join-Path $Root 'tools/template-lint-selftest.ps1'
    Assert-True (Test-Path -LiteralPath $selfTest) 'the template-lint self-test must exist'
    $exe = (Get-Process -Id $PID).Path
    $outFile = Join-Path ([IO.Path]::GetTempPath()) ("psq-tplst-{0}.out" -f [Guid]::NewGuid().ToString('N'))
    $proc = Start-Process -FilePath $exe `
        -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $selfTest) `
        -NoNewWindow -Wait -PassThru -RedirectStandardOutput $outFile
    $out = if (Test-Path -LiteralPath $outFile) { Get-Content -LiteralPath $outFile -Raw } else { '' }
    Remove-Item -LiteralPath $outFile -Force -ErrorAction SilentlyContinue
    Assert-Equal 0 $proc.ExitCode ('the template-lint self-test failed: ' + $out)
    Assert-Match 'FIRES/QUIET OK' $out 'the self-test must report per-fixture results'
}

Invoke-TestCase -Name 'No unguarded Measure-Object .Sum access' -Body {
    # Measure-Object returns $null (not an object) for an EMPTY pipeline, and
    # reading `.Sum` off $null throws under Set-StrictMode -Version Latest.
    # An empty result set is the normal case here - a clean spool directory, a
    # fleet where every host was unreachable - so the code must not depend on the
    # host running in lax mode. Nine such sites shipped before this rule existed.
    $bad = @()
    $files = @(Get-ChildItem -LiteralPath (Join-Path $Root 'src') -Recurse -Filter *.ps1 -File | Sort-Object FullName)
    foreach ($f in $files) {
        $lines = Get-Content -LiteralPath $f.FullName
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $ln = $lines[$i]
            if ($ln -match '^\s*#') { continue }
            if ($ln -match 'Measure-Object[^)]*\)\s*\.\s*Sum') { $bad += ("{0}:{1}" -f $f.Name, ($i + 1)) }
        }
    }
    Assert-True ($bad.Count -eq 0) $(if ($bad) { "unguarded .Sum after Measure-Object at: $($bad -join ', ')" } else { 'all sums are StrictMode-safe' })
}

Invoke-TestCase -Name 'No local variable collides with a typed parameter' -Body {
    # PowerShell variables are CASE-INSENSITIVE, so `$selfTest = Invoke-SelfTest`
    # inside a function that declares `[switch]$SelfTest` overwrites the parameter
    # and throws "Cannot convert value ... to SwitchParameter" - at the assignment,
    # not at the call, so the error points at an unrelated-looking object. This
    # exact defect shipped once and only showed up when the tool was loaded as a
    # module rather than dot-sourced.
    $scan = Join-Path $Root 'tools/Invoke-CollisionScan.ps1'
    Assert-True (Test-Path -LiteralPath $scan) 'the collision scanner must exist'
    # Run in a CHILD PROCESS and read the exit code from the process object.
    # $LASTEXITCODE is global mutable state: any earlier external command can
    # leave a value behind, which made this assertion flaky when the gate ran
    # nested inside Build.ps1.
    $exe = (Get-Process -Id $PID).Path
    $proc = Start-Process -FilePath $exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scan, '-SourceRoot', (Join-Path $Root 'src')) `
        -NoNewWindow -Wait -PassThru -RedirectStandardOutput (Join-Path ([IO.Path]::GetTempPath()) ("psq-scan-{0}.out" -f [Guid]::NewGuid().ToString('N')))
    Assert-Equal 0 $proc.ExitCode 'a parameter/local collision was detected (run tools/Invoke-CollisionScan.ps1 for details)'
}

Invoke-TestCase -Name 'The collision scanner detects a synthetic [switch] collision' -Body {
    # A detector that cannot fail is not a detector. Prove it fires on the exact
    # shape of the bug it exists to prevent.
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("psq-collide-{0}" -f [Guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $tmp -Force
    try {
        $bad = @'
function Test-Bad {
    [CmdletBinding()]
    param([switch]$SelfTest)
    $selfTest = Get-Date
    return $selfTest
}
'@
        Set-Content -LiteralPath (Join-Path $tmp 'Bad.ps1') -Value $bad -Encoding UTF8
        $exe = (Get-Process -Id $PID).Path
        $outFile = Join-Path ([IO.Path]::GetTempPath()) ("psq-scan-bad-{0}.out" -f [Guid]::NewGuid().ToString('N'))
        $proc = Start-Process -FilePath $exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $Root 'tools/Invoke-CollisionScan.ps1'), '-SourceRoot', $tmp) `
            -NoNewWindow -Wait -PassThru -RedirectStandardOutput $outFile
        $out = if (Test-Path -LiteralPath $outFile) { Get-Content -LiteralPath $outFile -Raw } else { '' }
        Remove-Item -LiteralPath $outFile -Force -ErrorAction SilentlyContinue
        Assert-Equal 1 $proc.ExitCode 'the scanner failed to flag a [switch] parameter being overwritten'
        Assert-Match 'COLLISION' $out 'the scanner must name the collision it found'
        Assert-Match 'SelfTest' $out 'the scanner must name the offending parameter'
    }
    finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

Invoke-TestCase -Name 'No .ContainsKey() on ordered dictionaries' -Body {
    $bad = @()
    foreach ($f in $sourceFiles) {
        $n = 0
        foreach ($line in (Get-Content -LiteralPath $f.FullName)) {
            $n++
            # $PSBoundParameters legitimately has ContainsKey; [ordered]@{} does not.
            if ($line -match '\.ContainsKey\(' -and $line -notmatch 'PSBoundParameters') { $bad += "$($f.Name):$n" }
        }
    }
    Assert-Equal 0 $bad.Count ("offenders: " + ($bad -join ', '))
}

Invoke-TestCase -Name 'No nested function declarations inside functions' -Body {
    $bad = @()
    foreach ($f in $sourceFiles) {
        $lines = Get-Content -LiteralPath $f.FullName
        $depth = 0; $inFunction = $false; $n = 0
        foreach ($line in $lines) {
            $n++
            if ($line -match '^\s*#\s*NOTE.*nested' -or $line -match "^\s*#") { continue }
            if (-not $inFunction -and $line -match '^function\s') { $inFunction = $true; $depth = 0 }
            if ($inFunction) {
                $depth += ([regex]::Matches($line, '\{')).Count
                $depth -= ([regex]::Matches($line, '\}')).Count
                if ($depth -gt 0 -and $line -match '^\s+function\s+[A-Z]') { $bad += "$($f.Name):$n $($line.Trim())" }
                if ($depth -le 0 -and $n -gt 1) { $inFunction = $false }
            }
        }
    }
    Assert-Equal 0 $bad.Count ("offenders: " + ($bad -join '; '))
}

Invoke-TestCase -Name 'No tabs anywhere in the sources' -Body {
    $bad = @()
    foreach ($f in $sourceFiles) {
        if ((Get-Content -LiteralPath $f.FullName -Raw) -match "`t") { $bad += $f.Name }
    }
    Assert-Equal 0 $bad.Count ("offenders: " + ($bad -join ', '))
}

Invoke-TestCase -Name 'Every file carries the author/license header or is generated' -Body {
    $missing = @()
    foreach ($f in $sourceFiles) {
        $head = (Get-Content -LiteralPath $f.FullName -TotalCount 40) -join "`n"
        if ($head -notmatch '#|<#') { $missing += $f.Name }
    }
    Assert-Equal 0 $missing.Count ("no header: " + ($missing -join ', '))
}

Invoke-TestCase -Name 'Value-returning helpers do not write to the host' -Body {
    # A stray Write-Host inside a helper that also emits a value corrupts the
    # caller's pipeline. The rule is deliberately narrow: it only flags functions
    # that BOTH declare an [OutputType] AND contain Write-Host, and it exempts
    # the rendering layer, whose entire job is host output.
    $suspect = @()
    foreach ($f in $sourceFiles) {
        if ($f.Name -match '03-Console|13-HelpText') { continue }
        $text = Get-Content -LiteralPath $f.FullName -Raw
        foreach ($m in [regex]::Matches($text, '(?ms)^function\s+([\w-]+).*?^\}')) {
            $name = $m.Groups[1].Value
            $body = $m.Value
            # The entry point renders progress and summaries to the host by
            # design; it is exempt, everything else that declares an output
            # contract must stay off the host.
            if ($name -eq 'Clear-PrintSpoolerQueue') { continue }
            if ($body -match '\[OutputType\(' -and $body -match 'Write-Host') { $suspect += "$($f.Name):$name" }
        }
    }
    Assert-Equal 0 $suspect.Count ("suspect functions: " + ($suspect -join ', '))
}

$lint = Get-TestSummary
if ($lint.Failed -gt 0) { $stageFailures++ }
Write-Host ("`n  Source lint: {0} passed, {1} failed" -f $lint.Passed, $lint.Failed) -ForegroundColor $(if ($lint.Failed -eq 0) { 'Green' } else { 'Red' })

# ---------------------------------------------------------------------------
# Stage 4: batch launcher lint
# ---------------------------------------------------------------------------
Write-Stage 'STAGE 4/9  Batch launcher lint'
Reset-TestCounters
Start-TestSuite 'Batch launcher'

$batPath = Join-Path $dist 'Clear-PrintSpoolerQueue.bat'
if (-not (Test-Path -LiteralPath $batPath)) { $batPath = Join-Path $Root 'bat/Clear-PrintSpoolerQueue.bat' }
$batLines = @(Get-Content -LiteralPath $batPath)
$batRaw = $batLines -join "`n"

Invoke-TestCase -Name 'Batch file exists and is non-trivial' -Body {
    Assert-GreaterThan $batLines.Count 200 'a full launcher is expected'
}

Invoke-TestCase -Name 'Starts with @echo off and setlocal' -Body {
    Assert-Match '^@echo off' $batLines[0]
    Assert-True ($batRaw -match 'setlocal EnableExtensions EnableDelayedExpansion')
}

Invoke-TestCase -Name 'Pure ASCII (no mojibake in the OEM code page)' -Body {
    $bytes = [IO.File]::ReadAllBytes($batPath)
    $nonAscii = @($bytes | Where-Object { $_ -gt 127 }).Count
    Assert-Equal 0 $nonAscii 'non-ASCII bytes found'
}

Invoke-TestCase -Name 'No tab characters' -Body {
    $tabs = @($batLines | Where-Object { $_ -match "`t" }).Count
    Assert-Equal 0 $tabs 'tabs found'
}

Invoke-TestCase -Name 'Every goto/call target has a label' -Body {
    $labels = @($batLines | Where-Object { $_ -match '^:[A-Za-z_]' } | ForEach-Object { ($_ -replace '^:', '' -replace '\s.*$', '').ToLowerInvariant() })
    $targets = @([regex]::Matches($batRaw, '(?i)\b(?:goto|call)\s+:?([A-Za-z_][A-Za-z0-9_]*)') | ForEach-Object { $_.Groups[1].Value.ToLowerInvariant() })
    $missing = @($targets | Select-Object -Unique | Where-Object { $_ -notin @('eof', 'finish') -and $labels -notcontains $_ })
    Assert-Equal 0 $missing.Count ("missing labels: " + ($missing -join ', '))
}

Invoke-TestCase -Name 'No duplicate labels' -Body {
    $labels = @($batLines | Where-Object { $_ -match '^:[A-Za-z_]' } | ForEach-Object { ($_ -replace '^:', '' -replace '\s.*$', '') })
    $dupes = @($labels | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
    Assert-Equal 0 $dupes.Count ("duplicates: " + ($dupes -join ', '))
}

Invoke-TestCase -Name 'Every argument clause shifts exactly as many times as it consumes' -Body {
    $bad = @()
    $n = 0
    foreach ($line in $batLines) {
        $n++
        if ($line -notmatch '^if /I "%ARG%"==') { continue }
        $consumesValue = $line -match '%~2'
        $shifts = ([regex]::Matches($line, '\bshift\b')).Count
        if ($consumesValue -and $shifts -lt 2) { $bad += "line ${n}: takes a value but shifts $shifts time(s)" }
        if (-not $consumesValue -and $shifts -lt 1) { $bad += "line ${n}: shifts $shifts time(s)" }
    }
    Assert-Equal 0 $bad.Count ($bad -join '; ')
}

Invoke-TestCase -Name 'No duplicate argument spellings' -Body {
    $args2 = @([regex]::Matches($batRaw, '(?im)^if /I "%ARG%"=="([^"]+)"') | ForEach-Object { $_.Groups[1].Value.ToLowerInvariant() })
    $dupes = @($args2 | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
    Assert-Equal 0 $dupes.Count ("duplicates: " + ($dupes -join ', '))
}

Invoke-TestCase -Name 'Every value-taking launcher switch is guarded against a missing value' -Body {
    # SAFETY, not tidiness - this is the launcher half of defect 39.
    #
    # An empty string is FALSY in PowerShell, so a blank selection filter is not
    # "match nothing": it is "no filter at all", and the run silently widens to
    # EVERY printer and every job. The launcher produced exactly that when an
    # operator typed `/printer` and forgot the value, forwarding -PrinterName "".
    #
    # Each value-taking switch must therefore either appear in the NEEDS_VALUE
    # guard list, or test for an empty %~2 inline (/help does, because its section
    # argument is genuinely optional).
    #
    # The point of asserting it here is the future: without this rule, adding a new
    # value-taking switch and forgetting to extend the guard ships a
    # scope-widening typo that no other check can see - the .bat never executes in
    # this test suite, because there is no cmd.exe to execute it.
    $guardMatch = [regex]::Match($batRaw, 'for %%S in \(([^)]*)\) do')
    Assert-True $guardMatch.Success 'the launcher must declare a NEEDS_VALUE guard list'
    $guard = @($guardMatch.Groups[1].Value -split '\s+' |
        Where-Object { $_ } |
        ForEach-Object { $_.ToLowerInvariant() })

    $unguarded = @()
    $valueTaking = 0
    foreach ($line in $batLines) {
        $m = [regex]::Match($line, '^if /I "%ARG%"=="([^"]+)".*%~2')
        if (-not $m.Success) { continue }
        $valueTaking++
        $sw = $m.Groups[1].Value.ToLowerInvariant()
        if ($guard -contains $sw) { continue }
        # An inline emptiness test is an equally valid guard.
        if ($line -match 'if "%~2"==""') { continue }
        $unguarded += $sw
    }

    Assert-GreaterThan $valueTaking 20 'the launcher should have many value-taking switches; the parse may have broken'
    Assert-Equal 0 $unguarded.Count $(if ($unguarded) {
            'value-taking switch(es) with no missing-value guard: ' + ($unguarded -join ', ')
        } else {
            "$valueTaking value-taking switches, all guarded ($($guard.Count) in NEEDS_VALUE)"
        })
}

Invoke-TestCase -Name 'The launcher guard list contains no switch that takes no value' -Body {
    # The other direction. A stale entry in NEEDS_VALUE would make the launcher
    # demand a value for a flag that takes none, so `/force extra` would be
    # rejected - or worse, an operator would learn to pass junk values to satisfy
    # a guard that should not exist.
    $guardMatch = [regex]::Match($batRaw, 'for %%S in \(([^)]*)\) do')
    Assert-True $guardMatch.Success 'the launcher must declare a NEEDS_VALUE guard list'
    $guard = @($guardMatch.Groups[1].Value -split '\s+' |
        Where-Object { $_ } |
        ForEach-Object { $_.ToLowerInvariant() })

    $valueTaking = @{}
    foreach ($line in $batLines) {
        $m = [regex]::Match($line, '^if /I "%ARG%"=="([^"]+)"')
        if (-not $m.Success) { continue }
        $sw = $m.Groups[1].Value.ToLowerInvariant()
        if (-not $valueTaking.ContainsKey($sw)) { $valueTaking[$sw] = $false }
        if ($line -match '%~2') { $valueTaking[$sw] = $true }
    }

    $stale = @($guard | Where-Object { -not $valueTaking.ContainsKey($_) -or -not $valueTaking[$_] })
    Assert-Equal 0 $stale.Count $(if ($stale) {
            'NEEDS_VALUE lists switch(es) that take no value: ' + ($stale -join ', ')
        } else {
            'all guarded switches really do take a value'
        })
}

Invoke-TestCase -Name 'Documented exit codes are all reachable in the code' -Body {
    $documented = @(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15)
    $found = @([regex]::Matches($batRaw, 'EXIT_CODE=(\d+)') | ForEach-Object { [int]$_.Groups[1].Value })
    foreach ($code in @(2, 3, 5, 10)) { Assert-Contains $found $code "exit code $code is never set" }
}

Invoke-TestCase -Name 'Version stamp matches the built script' -Body {
    $batVer = [regex]::Match($batRaw, 'set "SCRIPT_VERSION=([^"]+)"').Groups[1].Value
    $psVer = [regex]::Match((Get-Content -LiteralPath (Join-Path $dist 'Clear-PrintSpoolerQueue.ps1') -Raw), "\`$Script:ScriptVersion\s+=\s+'([^']+)'").Groups[1].Value
    Assert-Equal $psVer $batVer 'launcher and script must agree'
}

Invoke-TestCase -Name 'Never invokes the .bat recursively and never pauses unattended' -Body {
    Assert-Match 'FLAG_INTERACTIVE' $batRaw
    Assert-Match 'pause >nul' $batRaw
    Assert-True ($batRaw -match 'if "%FLAG_INTERACTIVE%"=="1"')
}

Invoke-TestCase -Name 'Delegates the cleanup to PowerShell (single destructive implementation)' -Body {
    Assert-Match '-ExecutionPolicy Bypass -File "%PS_SCRIPT%"' $batRaw
    # The launcher must not contain its own delete loop.
    Assert-NotMatch '(?im)^\s*del\s+/[fq]' $batRaw 'launcher performs its own deletion'
    Assert-NotMatch '(?im)^\s*net\s+stop\s+Spooler' $batRaw 'launcher stops the service itself'
}

$batSummary = Get-TestSummary
if ($batSummary.Failed -gt 0) { $stageFailures++ }
Write-Host ("`n  Batch lint: {0} passed, {1} failed" -f $batSummary.Passed, $batSummary.Failed) -ForegroundColor $(if ($batSummary.Failed -eq 0) { 'Green' } else { 'Red' })

# ---------------------------------------------------------------------------
# Stage 5: unit tests against the BUILT artifact
# ---------------------------------------------------------------------------
Write-Stage 'STAGE 5/9  Unit tests (against dist/Clear-PrintSpoolerQueue.ps1)'
Reset-TestCounters

$unitFiles = @(Get-ChildItem -LiteralPath (Join-Path $Root 'tests') -Filter 'Unit.*.ps1' -File | Sort-Object Name)
if ($UnitOnly -or $true) {
    . $artifact
    foreach ($uf in $unitFiles) {
        Write-Host ("`n  loading {0}" -f $uf.Name) -ForegroundColor DarkGray
        . $uf.FullName
    }
}
$unitSummary = Get-TestSummary
if ($unitSummary.Failed -gt 0) { $stageFailures++ }
Write-Host ''
if ($unitSummary.Failed -gt 0) {
    Write-Host '  Failed unit tests:' -ForegroundColor Red
    $unitSummary.Results | Where-Object { $_.Result -eq 'FAIL' } | ForEach-Object {
        Write-Host ("    [{0}] {1}`n        {2}" -f $_.Suite, $_.Name, $_.Detail) -ForegroundColor Red
    }
}
Write-Host ("  Unit tests: {0} passed, {1} failed" -f $unitSummary.Passed, $unitSummary.Failed) -ForegroundColor $(if ($unitSummary.Failed -eq 0) { 'Green' } else { 'Red' })

# ---------------------------------------------------------------------------
# Stage 6: module import test (the .psd1/.psm1 form is a separate code path from
# dot-sourcing: module scope, Export-ModuleMember and private-helper isolation
# all behave differently, and a defect here is invisible to every other stage)
# ---------------------------------------------------------------------------
$moduleExit = 0
if (-not $UnitOnly) {
    Write-Stage 'STAGE 6/9  Module import test'
    $moduleRunner = Join-Path $Root 'tools/module-test.ps1'
    if (Test-Path -LiteralPath $moduleRunner) {
        & (Get-Process -Id $PID).Path -NoProfile -ExecutionPolicy Bypass -File $moduleRunner
        $moduleExit = $LASTEXITCODE
        if ($moduleExit -ne 0) { $stageFailures++ }
    }
    else { Write-Host '   (no module-test.ps1; skipped)' -ForegroundColor Yellow }
}

# ---------------------------------------------------------------------------
# Stage 7: end-to-end smoke test (separate process, clean exit code)
# ---------------------------------------------------------------------------
$smokeExit = 0
if (-not $UnitOnly) {
    Write-Stage 'STAGE 7/9  End-to-end smoke test'
    $smokeArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $Root 'tools/smoke-test.ps1'), '-Root', $Root)
    if ($SkipNetwork) { $smokeArgs += '-SkipNetwork' }
    & (Get-Process -Id $PID).Path @smokeArgs
    $smokeExit = $LASTEXITCODE
    if ($smokeExit -ne 0) { $stageFailures++ }
}

# ---------------------------------------------------------------------------
# Stage 8: release artifact integrity - checksums, ZIP contents, version
# consistency across every shipped surface, documentation set, licence coherence.
# This is the gate between "the code is correct" and "the release is shippable".
# ---------------------------------------------------------------------------
if (-not $UnitOnly) {
    Write-Stage 'STAGE 8/9  Release artifact integrity'
    $verifier = Join-Path $Root 'tools/Verify-Release.ps1'
    if (Test-Path -LiteralPath $verifier) {
        & (Get-Process -Id $PID).Path -NoProfile -ExecutionPolicy Bypass -File $verifier
        if ($LASTEXITCODE -ne 0) { $stageFailures++ }
    }
    else { Write-Host '   (no Verify-Release.ps1; skipped)' -ForegroundColor Yellow }
}

# ---------------------------------------------------------------------------
# Stage 9: requirements traceability. Every other stage proves the code is
# internally CORRECT; this one proves it still does what the brief asked for.
#
# A large refactor can silently drop a documented capability - the gate would
# stay green while the product lost a feature nobody asserted by name. This stage
# exercises each requirement from the original brief against the built artifact:
# parameter metadata, exported cmdlets, implemented update-discovery strategies,
# rendered help sections, wiki page set, and live invocations.
# ---------------------------------------------------------------------------
if (-not $UnitOnly) {
    Write-Stage 'STAGE 9/9  Brief compliance (requirements traceability)'
    $brief = Join-Path $Root 'tools/Verify-Brief.ps1'
    if (Test-Path -LiteralPath $brief) {
        $briefArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $brief, '-Root', $Root)
        if ($SkipNetwork) { $briefArgs += '-SkipNetwork' }
        & (Get-Process -Id $PID).Path @briefArgs
        if ($LASTEXITCODE -ne 0) { $stageFailures++ }
    }
    else { Write-Host '   (no Verify-Brief.ps1; skipped)' -ForegroundColor Yellow }
}

# ---------------------------------------------------------------------------
$env:PSQ_LOG_PATH = $savedLogEnv
Remove-Item -LiteralPath $gateLogDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "`n$('=' * 78)" -ForegroundColor Cyan
if ($stageFailures -eq 0) {
    Write-Host '  QUALITY GATE PASSED' -ForegroundColor Green
}
else {
    Write-Host ("  QUALITY GATE FAILED ({0} stage(s))" -f $stageFailures) -ForegroundColor Red
}
Write-Host ('=' * 78) -ForegroundColor Cyan
exit $(if ($stageFailures -eq 0) { 0 } else { 1 })

<#
.SYNOPSIS
    Detects assignments to PowerShell AUTOMATIC variables that corrupt host state.

.DESCRIPTION
    `Invoke-CollisionScan.ps1` catches a local variable colliding with a function
    PARAMETER. This catches the sibling hazard: a local colliding with an AUTOMATIC
    variable the host maintains.

    These are worse than parameter collisions because they fail silently and at a
    distance:

      * `$args = @(...)` appears to work, then the host repopulates it and the
        value vanishes - or silently swallows the caller's own unbound arguments.
      * `$Matches = ...` destroys the result of the last `-match` for the caller.
      * `$Error = ...` wipes every error record the caller was about to inspect.
      * `$_ = ...` inside a pipeline block corrupts the item being processed.

    Two instances of the `$args` form shipped in this project before this scanner
    existed: one in the CI release job (`gh @args`) and one in the production
    `New-ChildArgumentList` helper.

    Deliberately NOT reported, because both are correct and pervasive:

      * `$null = <expression>` - the idiomatic way to discard pipeline output.
        `$null` accepts assignment by design; it is not a hazard.
      * The `*Preference` variables and `$ErrorActionPreference` - these exist to
        be assigned. Setting them at script scope is standard practice.

.PARAMETER SourceRoots
    Directories to scan. Defaults to src/, tools/, tests/ and the CI workflow.

.EXAMPLE
    ./tools/Invoke-AutomaticVariableScan.ps1
.EXAMPLE
    ./tools/Invoke-AutomaticVariableScan.ps1 -Quiet
#>

[CmdletBinding()]
param(
    # Explicit roots to scan. When omitted the project's own src/, tools/, tests/
    # and CI workflow are used. A caller-supplied root is what makes the detector
    # TESTABLE: the lint suite points it at a fixture containing a known defect
    # and asserts the scanner fires. A detector that can only ever look at the
    # real tree cannot be proven to work.
    [AllowNull()][string[]]$SourceRoots,

    # Skip the CI workflow scan (it lives outside any source root).
    [switch]$SkipWorkflow,

    [switch]$Quiet
)

$root = Split-Path -Parent $PSScriptRoot

# Assignment corrupts host state or the caller's expectations.
$dangerous = @(
    'args', 'input', 'matches', 'error', 'psitem', '_', 'this', 'event', 'eventargs',
    'eventsubscriber', 'stacktrace', 'conflicterror', 'psboundparameters',
    'myinvocation', 'pscmdlet', 'switch', 'foreach', 'executioncontext',
    'psdebugcontext', 'nestedpromptlevel', 'lastexitcode', 'ofs', 'outputencoding'
)

# Assignment throws immediately, so it is self-reporting rather than silent - but
# still a defect worth failing the build on.
$readOnly = @(
    'psversiontable', 'pid', 'host', 'home', 'pwd', 'shellid', 'psculture',
    'psuiculture', 'iswindows', 'islinux', 'ismacos', 'true', 'false',
    'psscriptroot'
)

# Correct and pervasive; excluded on purpose (see .DESCRIPTION).
$allowed = @('null', 'erroractionpreference', 'warningpreference', 'verbosepreference',
    'debugpreference', 'informationpreference', 'progresspreference', 'confirmPreference',
    'errorview', 'pscommandpath', 'psedition')

$targets = New-Object System.Collections.Generic.List[string]
$rootsToScan = if ($SourceRoots -and $SourceRoots.Count -gt 0) { $SourceRoots } else { @('src', 'tools', 'tests' | ForEach-Object { Join-Path $root $_ }) }
foreach ($path in $rootsToScan) {
    if (-not (Test-Path -LiteralPath $path)) { continue }
    if (Test-Path -LiteralPath $path -PathType Leaf) {
        if ($path -like '*.ps1') { $targets.Add((Resolve-Path -LiteralPath $path).Path) }
        continue
    }
    Get-ChildItem -LiteralPath $path -Recurse -Filter *.ps1 -File |
        ForEach-Object { $targets.Add($_.FullName) }
}

$hits = New-Object System.Collections.Generic.List[string]
$scanned = 0

foreach ($file in $targets) {
    $tokens = $null; $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($file, [ref]$tokens, [ref]$errors)
    if (-not $ast) { continue }
    $scanned++
    $name = Split-Path -Leaf $file

    foreach ($a in @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true))) {
        $m = [regex]::Match($a.Left.Extent.Text, '^\$(\w+)$')
        if (-not $m.Success) { continue }
        $raw = $m.Groups[1].Value
        $lower = $raw.ToLowerInvariant()
        if ($allowed -contains $lower) { continue }

        if ($dangerous -contains $lower) {
            $rhs = $a.Right.Extent.Text
            $hits.Add(("{0}:{1}  WRITES AUTOMATIC VARIABLE  `${2} = {3}" -f $name,
                $a.Extent.StartLineNumber, $raw, $rhs.Substring(0, [Math]::Min(48, $rhs.Length))))
        }
        elseif ($readOnly -contains $lower) {
            $hits.Add(("{0}:{1}  WRITES READ-ONLY VARIABLE    `${2}" -f $name,
                $a.Extent.StartLineNumber, $raw))
        }
    }

    # Splatting an automatic variable: legal syntax, but it means a hand-built
    # list was stored in a name the host owns. That is how the CI defect shipped.
    foreach ($t in $tokens) {
        if ($t.Kind -eq 'Splat' -and ($t.Text -in @('@args', '@input', '@matches', '@error'))) {
            $hits.Add(("{0}:{1}  SPLATS AN AUTOMATIC VARIABLE  {2}" -f $name, $t.Extent.StartLineNumber, $t.Text))
        }
    }
}

# The CI workflow is YAML with embedded PowerShell, invisible to the AST. Scan it
# textually for the exact pattern that shipped, skipping comments.
$ci = Join-Path $root '.github/workflows/ci.yml'
if ((-not $SkipWorkflow) -and (-not $SourceRoots) -and (Test-Path -LiteralPath $ci)) {
    $scanned++
    $n = 0
    foreach ($line in @(Get-Content -LiteralPath $ci)) {
        $n++
        if ($line -match '^\s*#') { continue }
        foreach ($v in ($dangerous + $readOnly)) {
            if ($line -match ('(?<![\w$])\$' + [regex]::Escape($v) + '\s*(=[^=]|\+=|-=|\*=|/=)')) {
                $hits.Add(("ci.yml:{0}  WRITES AUTOMATIC VARIABLE  `${1}" -f $n, $v))
            }
        }
        foreach ($v in @('args', 'input', 'matches')) {
            if ($line -match ('(?<![\w$])@' + [regex]::Escape($v) + '(?![\w])')) {
                $hits.Add(("ci.yml:{0}  SPLATS AN AUTOMATIC VARIABLE  @{1}" -f $n, $v))
            }
        }
    }
}

if (-not $Quiet) {
    Write-Host ("Scanned {0} file(s) for automatic-variable writes." -f $scanned) -ForegroundColor Cyan
    Write-Host ("Excluded by design: `$null = ... (output discard) and the *Preference variables.") -ForegroundColor DarkGray
    Write-Host ''
}
if ($hits.Count -eq 0) {
    if (-not $Quiet) { Write-Host 'No automatic variable is assigned or splatted anywhere.' -ForegroundColor Green }
    exit 0
}
if (-not $Quiet) {
    Write-Host ("{0} FINDING(S):" -f $hits.Count) -ForegroundColor Red
    $hits | ForEach-Object { Write-Host ("  {0}" -f $_) -ForegroundColor Yellow }
}
exit 1

<#
.SYNOPSIS
    Detects PowerShell 7+-only constructs in code that must also run on Windows
    PowerShell 5.1.

.DESCRIPTION
    The product's declared floor is Windows PowerShell 5.1. Most PS7-only syntax
    is a hard PARSE error there, which is loud and therefore survivable. Two
    categories are not:

      * `$IsWindows` / `$IsLinux` / `$IsMacOS` are simply $null on 5.1, so a
        branch written as `if ($IsWindows)` silently takes the wrong path instead
        of failing. This is the dangerous one.
      * `ForEach-Object -Parallel` is legitimate here - the tool genuinely uses it
        for fleet dispatch - but only behind a runtime version guard, with a
        runspace-pool fallback for 5.1. An UNGUARDED call parses fine on 5.1 and
        then fails at bind time, on whatever machine happens to be old.

    Detection is AST/token based rather than line-regex based, which matters for
    two reasons: comments and strings never produce the tokens being matched, so
    the word "-Parallel" inside help text is not a finding; and `$IsWindows` as a
    hashtable KEY (`IsWindows = ...`) or as a string passed to `Get-Variable` is
    correctly distinguished from a real variable reference.

.PARAMETER SourceRoots
    Directories or files to scan. Defaults to src/ - the only tree that ships.
    Supplying roots explicitly is what makes the detector testable: the lint suite
    points it at fixtures and asserts it both fires and stays quiet.

.PARAMETER IncludeArtifacts
    Also scan dist/Clear-PrintSpoolerQueue.ps1 and the module .psm1. The gate uses
    this so the SHIPPED bytes are verified, not just the sources they came from.

.PARAMETER Quiet
    Suppress progress output; exit code only.

.EXAMPLE
    ./tools/Invoke-PS51CompatibilityScan.ps1
.EXAMPLE
    ./tools/Invoke-PS51CompatibilityScan.ps1 -SourceRoots $fixture -Quiet
#>

[CmdletBinding()]
param(
    [AllowNull()][string[]]$SourceRoots,
    [switch]$IncludeArtifacts,
    [switch]$Quiet
)

$root = Split-Path -Parent $PSScriptRoot

# Resolve the file list.
$files = New-Object System.Collections.Generic.List[string]
$rootsToScan = if ($SourceRoots -and $SourceRoots.Count -gt 0) { $SourceRoots } else { @(Join-Path $root 'src') }
foreach ($r in $rootsToScan) {
    if (-not (Test-Path -LiteralPath $r)) { continue }
    if (Test-Path -LiteralPath $r -PathType Leaf) {
        if ($r -like '*.ps1') { $files.Add((Resolve-Path -LiteralPath $r).Path) }
        continue
    }
    Get-ChildItem -LiteralPath $r -Recurse -Filter *.ps1 -File | ForEach-Object { $files.Add($_.FullName) }
}
if ($IncludeArtifacts) {
    foreach ($a in @('dist/Clear-PrintSpoolerQueue.ps1', 'dist/PrintSpoolerCleanup/PrintSpoolerCleanup.psm1')) {
        $p = Join-Path $root $a
        if (Test-Path -LiteralPath $p) { $files.Add($p) }
    }
}

# Parameters that do not exist on 5.1, keyed by the command they belong to.
$ps7Params = @{
    'ConvertFrom-Json' = @('AsHashtable', 'Depth')
    'Join-Path'        = @('AdditionalChildPath')
    'ForEach-Object'   = @('ThrottleLimit')
    'Where-Object'     = @()
}
# Cmdlets that do not exist on 5.1 at all.
$ps7Cmdlets = @('Get-Error', 'Test-Json', 'ConvertTo-Json', 'Join-Path')  # ConvertTo-Json/Join-Path exist; filtered below
$ps7CmdletsReal = @('Get-Error', 'Test-Json', 'Get-Uptime', 'Clear-RecycleBin')

$findings = New-Object System.Collections.Generic.List[string]
$scanned = 0

function Test-InsideVersionGuard {
    <#
    .SYNOPSIS
        True when the given node sits inside an `if` whose CONDITION performs a
        PowerShell version check.
    .NOTES
        IfStatementAst has no Conditions() method in the shipped AST, so the
        condition pipeline is read from Clauses[0].Item1. Falling back to the whole
        Extent.Text would be wrong: a large if-body that merely mentions the guard
        helper would look guarded when it is not.
    #>
    param([Parameter(Mandatory)][AllowNull()]$Node)
    $p = $Node
    while ($p) {
        if ($p -is [System.Management.Automation.Language.IfStatementAst]) {
            $condText = $null
            try {
                $clauses = @($p.Clauses)
                if ($clauses.Count -gt 0 -and $clauses[0].Item1) { $condText = $clauses[0].Item1.Extent.Text }
            }
            catch { $condText = $null }
            if (-not $condText) { $condText = $p.Extent.Text }
            if ($condText -match 'Test-PowerShellSevenFeatures|PSVersionTable|\$PSVersion') { return $true }
            # Not guarded by THIS if, but an enclosing if may guard it.
        }
        $p = $p.Parent
    }
    return $false
}

foreach ($file in $files) {
    $tokens = $null; $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($file, [ref]$tokens, [ref]$errors)
    if (-not $ast) { continue }
    $scanned++
    $name = Split-Path -Leaf $file

    # ---- 1. Token-level PS7-only operators -------------------------------
    # Matched on token TEXT, not token Kind: Kind names differ across PowerShell
    # versions, and token text is stable. Comments and string literals never emit
    # these tokens, so documentation mentioning "??" is not a finding.
    foreach ($tok in $tokens) {
        switch ($tok.Text) {
            '??' {
                $findings.Add(("{0}:{1}  null-coalescing operator '??' is PowerShell 7+" -f $name, $tok.Extent.StartLineNumber))
            }
            '??=' {
                $findings.Add(("{0}:{1}  null-coalescing assignment '??=' is PowerShell 7+" -f $name, $tok.Extent.StartLineNumber))
            }
            '?' {
                if ($tok.Kind -match 'Question') {
                    $findings.Add(("{0}:{1}  ternary operator '? :' is PowerShell 7+" -f $name, $tok.Extent.StartLineNumber))
                }
            }
        }
        # Null-conditional member/index access tokenizes as a Variable whose text
        # ENDS WITH '?', followed by Dot or LBracket. There is no dedicated kind.
        if ($tok.Kind -eq 'Variable' -and $tok.Text.EndsWith('?')) {
            $findings.Add(("{0}:{1}  null-conditional access '{2}' is PowerShell 7+" -f $name, $tok.Extent.StartLineNumber, $tok.Text))
        }
    }

    # ---- 2. clean{} block (PowerShell 7.4+) -------------------------------
    foreach ($fn in $ast.FindAll({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
        if ($fn.Body -and $fn.Body.CleanBlock) {
            $findings.Add(("{0}:{1}  function '{2}' has a clean{{}} block (PowerShell 7.4+)" -f $name, $fn.Extent.StartLineNumber, $fn.Name))
        }
    }

    # ---- 3. $IsWindows / $IsLinux / $IsMacOS -------------------------------
    # The silent-failure case: these are $null on 5.1, so `if ($IsWindows)` takes
    # the wrong branch instead of throwing. Only real variable references count;
    # a hashtable key or a string argument is correctly ignored by the AST.
    foreach ($v in $ast.FindAll({ param($x) $x -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)) {
        if ($v.VariablePath.UserPath -match '^(IsWindows|IsLinux|IsMacOS)$') {
            $findings.Add(("{0}:{1}  automatic variable '{2}' does not exist on 5.1 (it is `$null there, so the branch fails SILENTLY) - use Test-RunningOnWindows" -f $name, $v.Extent.StartLineNumber, $v.Extent.Text))
        }
    }

    # ---- 4. Commands and parameters that do not exist on 5.1 ---------------
    foreach ($cmd in $ast.FindAll({ param($x) $x -is [System.Management.Automation.Language.CommandAst] }, $true)) {
        $elems = @($cmd.CommandElements)
        if ($elems.Count -eq 0) { continue }
        $cmdName = $elems[0].Extent.Text

        if ($ps7CmdletsReal -contains $cmdName) {
            $findings.Add(("{0}:{1}  cmdlet '{2}' does not exist on PowerShell 5.1" -f $name, $cmd.Extent.StartLineNumber, $cmdName))
        }

        if (-not $ps7Params.ContainsKey($cmdName)) { continue }
        $bad = @($ps7Params[$cmdName])
        foreach ($el in $elems) {
            if ($el -isnot [System.Management.Automation.Language.CommandParameterAst]) { continue }
            # ThrottleLimit is legitimate when the whole call is version-guarded
            # (it is only ever paired with -Parallel here).
            if ($cmdName -eq 'ForEach-Object' -and $el.ParameterName -eq 'ThrottleLimit') {
                if (Test-InsideVersionGuard -Node $cmd) { continue }
            }
            if ($bad -contains $el.ParameterName) {
                $findings.Add(("{0}:{1}  {2} -{3} is not available on PowerShell 5.1" -f $name, $cmd.Extent.StartLineNumber, $cmdName, $el.ParameterName))
            }
        }

        # -Parallel specifically: allowed, but ONLY behind a version guard.
        foreach ($el in $elems) {
            if ($el -is [System.Management.Automation.Language.CommandParameterAst] -and $el.ParameterName -eq 'Parallel') {
                if (-not (Test-InsideVersionGuard -Node $cmd)) {
                    $findings.Add(("{0}:{1}  ForEach-Object -Parallel is UNGUARDED - on 5.1 this parses and then fails at bind time; wrap it in Test-PowerShellSevenFeatures" -f $name, $cmd.Extent.StartLineNumber))
                }
            }
        }
    }
}

if (-not $Quiet) {
    Write-Host ("Scanned {0} file(s) for PowerShell 7+-only constructs." -f $scanned) -ForegroundColor Cyan
    Write-Host 'Detection is AST/token based: comments, help text and hashtable keys are not findings.' -ForegroundColor DarkGray
    Write-Host ''
}
if ($findings.Count -eq 0) {
    if (-not $Quiet) { Write-Host 'No PowerShell 5.1 incompatibilities found.' -ForegroundColor Green }
    exit 0
}
if (-not $Quiet) {
    Write-Host ("{0} FINDING(S):" -f $findings.Count) -ForegroundColor Red
    $findings | ForEach-Object { Write-Host ("  {0}" -f $_) -ForegroundColor Yellow }
}
exit 1

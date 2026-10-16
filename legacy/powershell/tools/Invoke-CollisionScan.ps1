<#
.SYNOPSIS
    Detects the most dangerous silent bug class in PowerShell: a local variable that
    collides, case-insensitively, with a typed parameter of the enclosing function.

.DESCRIPTION
    PowerShell variables are CASE-INSENSITIVE, so in

        param([switch]$SelfTest)
        $selfTest = Invoke-Something      # <- this IS the parameter

    the assignment overwrites the switch parameter. For a `[switch]` this always
    throws "Cannot convert value ... to type SwitchParameter" - and because the
    throw happens at the assignment, not at the call, the error surfaces deep
    inside the function while pointing at an unrelated-looking object.

    Reassigning a `[string]`/`[int]` parameter to normalise its input is normal,
    legitimate PowerShell and is NOT reported. Only assignments that cannot
    possibly satisfy the declared parameter type are flagged:

      * any assignment to a [switch] parameter (a switch can only ever hold a
        SwitchParameter, so `$x = <anything else>` is always fatal)
      * an assignment of an un-cast command result to a value-typed parameter
        ([int], [long], [bool], [double], [TimeSpan], [DateTime], [Version])

.NOTES
    Only the function's OWN param block is considered. Nested scriptblocks (for
    example the worker passed to ForEach-Object -Parallel or AddScript) declare
    their own parameters and initialise them deliberately; treating those as
    collisions produces pure noise.
#>

[CmdletBinding()]
param(
    [string]$SourceRoot,
    [switch]$Quiet
)

$root = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($SourceRoot)) { $SourceRoot = Join-Path $root 'src' }

$files = @(Get-ChildItem -LiteralPath $SourceRoot -Recurse -Filter *.ps1 -File | Sort-Object FullName)
$collisions = New-Object System.Collections.Generic.List[string]
$scanned = 0

# Value types whose parameters cannot accept an arbitrary un-cast expression.
$valueTypes = @('int', 'int32', 'long', 'int64', 'bool', 'boolean', 'double', 'single',
    'decimal', 'timespan', 'datetime', 'version', 'nullable', 'switch', 'switchparameter')

foreach ($f in $files) {
    $tokens = $null; $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$tokens, [ref]$errors)
    if (-not $ast) { continue }

    $funcs = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true))
    foreach ($fn in $funcs) {
        # The function's OWN param block only.
        $body = $fn.Body
        if (-not $body -or -not $body.ParamBlock) { continue }
        $ownParams = @($body.ParamBlock.Parameters)
        if ($ownParams.Count -eq 0) { continue }
        $scanned++

        # Map lower-cased parameter name -> declared type name (or '' when untyped).
        $paramTypes = @{}
        foreach ($p in $ownParams) {
            $nm = $p.Name.VariablePath.UserPath
            if (-not $nm) { continue }
            $typeName = ''
            if ($p.StaticType) { $typeName = $p.StaticType.Name.ToLowerInvariant() }
            $paramTypes[$nm.ToLowerInvariant()] = $typeName
        }

        # Assignments in the body, EXCLUDING those inside nested scriptblocks that
        # declare their own param block.
        $assigns = @($body.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true))
        foreach ($a in $assigns) {
            $m = [regex]::Match($a.Left.Extent.Text, '^\$(\w+)$')
            if (-not $m.Success) { continue }
            $varName = $m.Groups[1].Value
            $lower = $varName.ToLowerInvariant()
            if (-not $paramTypes.ContainsKey($lower)) { continue }

            # Skip assignments that live inside a nested scriptblock with its own
            # param block - those initialise the inner block's parameters.
            $nested = $false
            $node = $a.Parent
            while ($node -and ($node -ne $body)) {
                if (($node -is [System.Management.Automation.Language.ScriptBlockExpressionAst]) -and
                    $node.ScriptBlock.ParamBlock) { $nested = $true; break }
                if (($node -is [System.Management.Automation.Language.ScriptBlockExpressionAst]) -and
                    $node.ScriptBlock.Ast.Body.ParamBlock) { $nested = $true; break }
                $node = $node.Parent
            }
            if ($nested) { continue }

            $declaredType = $paramTypes[$lower]
            $right = $a.Right.Extent.Text.Trim()

            # Re-reading the same parameter is harmless.
            if ($right -match ('^\$' + [regex]::Escape($varName) + '$')) { continue }

            $isSwitch = ($declaredType -eq 'switchparameter') -or ($declaredType -eq 'switch')
            if ($isSwitch) {
                # A switch parameter can never hold anything but a SwitchParameter.
                $collisions.Add(("{0}:{1} in '{2}': the [switch] parameter '{3}' is overwritten by  {4} = {5}" -f `
                    $f.Name, $a.Extent.StartLineNumber, $fn.Name, $varName, $a.Left.Extent.Text, $right))
                continue
            }

            if ($declaredType -and ($valueTypes -contains $declaredType)) {
                # An explicit cast or a literal is safe; a bare command call is not.
                # Honour .NET type aliases: an `[int]` cast satisfies an `[int32]`
                # parameter, `[bool]` satisfies `[boolean]`, `[long]` satisfies `[int64]`.
                $aliases = @{
                    'int32' = @('int', 'int32'); 'int64' = @('long', 'int64')
                    'boolean' = @('bool', 'boolean'); 'double' = @('double', 'single')
                    'timespan' = @('timespan'); 'datetime' = @('datetime'); 'version' = @('version')
                    'decimal' = @('decimal'); 'single' = @('single', 'float')
                }
                $accepted = if ($aliases.ContainsKey($declaredType)) { $aliases[$declaredType] } else { @($declaredType) }
                $hasCast = $false
                foreach ($al in $accepted) { if ($right.StartsWith('[' + $al)) { $hasCast = $true; break } }
                if (-not $hasCast) {
                    # [Math]::Max/Min/Abs/Round/Clamp on a numeric parameter is a
                    # deliberate, type-safe narrowing - not a collision.
                    if ($right -match '^\[Math\]::(Max|Min|Abs|Round|Ceiling|Floor|Clamp|Truncate)\(') { $hasCast = $true }
                }
                $isLiteral = ($right -match '^[-+]?[0-9.]+$') -or ($right -eq '$true') -or ($right -eq '$false')
                if (-not $hasCast -and -not $isLiteral) {
                    $collisions.Add(("{0}:{1} in '{2}': the [{3}] parameter '{4}' receives an un-cast expression:  {5}" -f `
                        $f.Name, $a.Extent.StartLineNumber, $fn.Name, $declaredType, $varName, $right))
                }
            }
        }
    }
}

if (-not $Quiet) {
    Write-Host ("Scanned {0} function(s) with a param block across {1} file(s)." -f $scanned, $files.Count) -ForegroundColor Cyan
    Write-Host ''
}
if ($collisions.Count -eq 0) {
    if (-not $Quiet) { Write-Host 'No parameter/local variable collisions found.' -ForegroundColor Green }
    exit 0
}
if (-not $Quiet) {
    Write-Host ("{0} COLLISION(S):" -f $collisions.Count) -ForegroundColor Red
    $collisions | ForEach-Object { Write-Host ("  {0}" -f $_) -ForegroundColor Yellow }
}
exit 1

<#
.SYNOPSIS
    Parse-validates the PowerShell embedded inside YAML workflow files.

.DESCRIPTION
    tools/Invoke-Validation.ps1 parses every *.ps1 with the real PowerShell
    parser. But the CI workflow embeds substantial PowerShell inside `run: |`
    blocks, and YAML is opaque to that validation. A syntax error there is
    invisible locally and only surfaces when the workflow runs on GitHub -
    after the push, in front of everyone, burning a runner.

    This extracts every `run:` block from .github/workflows/*.yml, reconstructs it
    with the block indentation stripped, and parses it with
    [System.Management.Automation.Language.Parser]. It also flags the constructs
    that are legal PowerShell but wrong inside a workflow:

      * assignment to an automatic variable ($args and friends), which the
        project's own scanner forbids in .ps1 files but could not previously see
        inside YAML
      * `exit 0` at the end of a step, which masks a failure from the runner

.PARAMETER Quiet
    Suppress progress output; exit code only.

.EXAMPLE
    ./tools/Invoke-WorkflowLint.ps1
#>

[CmdletBinding()]
param(
    [switch]$Quiet
)

$root = Split-Path -Parent $PSScriptRoot
$wfDir = Join-Path $root '.github/workflows'
if (-not (Test-Path -LiteralPath $wfDir)) {
    if (-not $Quiet) { Write-Host "No workflow directory at $wfDir" -ForegroundColor Yellow }
    exit 0
}

# Automatic variables: assigning one is the trap documented in the wiki
# (trap 18). The .ps1 scanner cannot see inside YAML, so it is checked here too.
$automatic = @(
    'args', 'input', 'matches', 'error', 'psitem', '_', 'this', 'event', 'eventargs',
    'eventsubscriber', 'stacktrace', 'conflicterror', 'psboundparameters',
    'myinvocation', 'pscmdlet', 'switch', 'foreach', 'executioncontext',
    'psdebugcontext', 'nestedpromptlevel', 'lastexitcode', 'ofs', 'outputencoding',
    'psversiontable', 'pid', 'host', 'home', 'pwd', 'shellid', 'true', 'false'
)
# Assigned by design and correct; excluded.
$allowed = @('null', 'erroractionpreference', 'warningpreference', 'verbosepreference',
    'debugpreference', 'informationpreference', 'progresspreference', 'confirmpreference',
    'errorview', 'pscommandpath', 'psedition')

$fail = 0
$pass = 0
$blocks = 0

function Report {
    param([string]$N, [bool]$Ok, [string]$D = '')
    if ($Ok) { $script:pass++; if (-not $Quiet) { Write-Host ("  PASS  {0,-58} {1}" -f $N, $D) -ForegroundColor Green } }
    else { $script:fail++; Write-Host ("  FAIL  {0,-58} {1}" -f $N, $D) -ForegroundColor Red }
}

foreach ($wf in @(Get-ChildItem -LiteralPath $wfDir -Filter '*.y*ml' -File | Sort-Object Name)) {
    $lines = @(Get-Content -LiteralPath $wf.FullName)
    if (-not $Quiet) { Write-Host ''; Write-Host ("--- {0} ({1} lines) ---" -f $wf.Name, $lines.Count) -ForegroundColor Cyan }

    # Walk the file collecting `run: |` (and `run: >`) literal blocks. A block
    # ends at the first line that is less-indented than the block body and is not
    # blank.
    $i = 0
    while ($i -lt $lines.Count) {
        $line = $lines[$i]
        if ($line -notmatch '^(\s*)(?:-\s*)?run:\s*[|>][-+]?\s*$') { $i++; continue }

        $stepLine = $i + 1
        $keyIndent = $Matches[1].Length
        # Find the step name by scanning backwards for `- name:`.
        $stepName = "(line $stepLine)"
        for ($b = $i; $b -ge 0 -and $b -gt ($i - 40); $b--) {
            if ($lines[$b] -match '^\s*-\s*name:\s*(.+?)\s*$') { $stepName = $Matches[1].Trim('"', "'"); break }
        }

        $body = New-Object System.Collections.Generic.List[string]
        $j = $i + 1
        $bodyIndent = -1
        while ($j -lt $lines.Count) {
            $bl = $lines[$j]
            if ($bl -match '^\s*$') { $body.Add(''); $j++; continue }
            $ind = ([regex]::Match($bl, '^\s*')).Length
            if ($bodyIndent -lt 0) { $bodyIndent = $ind }
            # A line at or before the `run:` key's indentation ends the block.
            if ($ind -le $keyIndent) { break }
            if ($ind -lt $bodyIndent) { break }
            $body.Add($bl.Substring([Math]::Min($bodyIndent, $bl.Length)))
            $j++
        }

        $blocks++
        $text = ($body -join "`n")
        if ([string]::IsNullOrWhiteSpace($text)) { $i = $j; continue }

        # 1. It must PARSE. This is the whole point: a YAML-embedded syntax error
        #    is otherwise discovered only on a GitHub runner.
        $tokens = $null; $errors = $null
        $null = [System.Management.Automation.Language.Parser]::ParseInput($text, [ref]$tokens, [ref]$errors)
        $label = ("{0} :: {1}" -f $wf.Name, $stepName)
        if ($errors -and $errors.Count -gt 0) {
            Report "parses: $label" $false ($errors[0].Message + ' @line ' + $errors[0].Extent.StartLineNumber)
        }
        else {
            Report "parses: $label" $true ("{0} lines" -f $body.Count)
        }

        # 2. No automatic-variable assignment (trap 18), skipping comments.
        $n = 0
        foreach ($bl in $body) {
            $n++
            if ($bl -match '^\s*#') { continue }
            foreach ($v in $automatic) {
                if ($allowed -contains $v) { continue }
                if ($bl -match ('(?<![\w$])\$' + [regex]::Escape($v) + '\s*(=[^=]|\+=|-=|\*=|/=)')) {
                    Report "no automatic-variable write: $label" $false ("`${0} at block line {1}" -f $v, $n)
                }
                if ($bl -match ('(?<![\w$])@' + [regex]::Escape($v) + '(?![\w])') -and ($v -in @('args','input','matches'))) {
                    Report "no automatic-variable splat: $label" $false ("@{0} at block line {1}" -f $v, $n)
                }
            }
        }

        # 3. A trailing `exit 0` hides a failure from the runner: the step goes
        #    green even when the command above it failed.
        $nonBlank = @($body | Where-Object { $_ -notmatch '^\s*$' -and $_ -notmatch '^\s*#' })
        if ($nonBlank.Count -gt 0 -and $nonBlank[-1] -match '^\s*exit\s+0\s*$') {
            Report "no failure-masking exit 0: $label" $false 'last statement is `exit 0`'
        }

        $i = $j
    }
}

if (-not $Quiet) {
    Write-Host ''
    Write-Host ('=' * 78) -ForegroundColor Cyan
    Write-Host ("  WORKFLOW LINT: {0} run-block(s) checked, {1} passed, {2} failed" -f $blocks, $pass, $fail) -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
    Write-Host ('=' * 78) -ForegroundColor Cyan
}
exit $(if ($fail -eq 0) { 0 } else { 1 })

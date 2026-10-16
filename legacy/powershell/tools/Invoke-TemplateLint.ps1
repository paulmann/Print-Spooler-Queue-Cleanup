<#
.SYNOPSIS
    Validates GitHub issue-form templates and YAML escape safety, without a YAML
    parser dependency.

.DESCRIPTION
    Two classes of defect here fail SILENTLY and are invisible to every other
    check in the gate:

      1. An invalid escape sequence in a DOUBLE-QUOTED YAML scalar. `".\Clear-..."`
         is a hard parse error ("unknown escape character 'C'"), and GitHub then
         renders "There was an error loading this form" instead of the content.
         Nothing fails at build time; the form is simply broken for every user.

      2. An issue-form body element that does not match GitHub's schema. A
         `markdown` element carries its text in `attributes.value`; every other
         element type requires `attributes.label`. A wrong or missing key means
         the element renders empty or the whole form is rejected.

    The escape check must be BLOCK-SCALAR AWARE, and this is where a naive
    line-scanner produces pure noise: inside a `run: |` block the content is a
    literal string, so the PowerShell regexes this project embeds there
    (`"\`$Script:ScriptVersion\s+=\s+'([^']+)'"`) contain backslashes that are NOT
    YAML escapes and must not be reported. A scanner that ignores block scalars
    flags six false positives in ci.yml and gets disabled.

    The project's rule is native utilities only - no powershell-yaml, no Pester -
    so this is a targeted structural validator rather than a full YAML parse. It
    checks the properties that actually break rendering.

.PARAMETER Root
    Repository root to inspect. Defaults to this project. Supplying a fixture
    root is how the self-test proves the linter actually fails on bad input.

.PARAMETER Quiet
    Suppress progress output; exit code only.

.EXAMPLE
    ./tools/Invoke-TemplateLint.ps1
#>

[CmdletBinding()]
param(
    # Repository root to inspect. Defaults to this project. An explicit root is
    # what makes the linter TESTABLE: the self-test points it at a fixture tree
    # containing known-bad templates and asserts it fails. Without this the
    # linter can only ever look at the real tree, so "it passed" would mean
    # nothing - a detector that cannot fail is not a detector.
    [AllowNull()][string]$Root,

    [switch]$Quiet
)

if ([string]::IsNullOrWhiteSpace($Root)) { $Root = Split-Path -Parent $PSScriptRoot }
$root = $Root

# YAML 1.2 double-quote escape set. Anything else after a backslash is a parse
# error, not a stylistic choice.
$validEscapes = @('0', 'a', 'b', 't', 'n', 'v', 'f', 'r', 'e', '"', '/', '\',
    'N', '_', 'L', 'P', 'x', 'u', 'U', ' ', "`t")

# GitHub issue-form element types and where each keeps its text.
$markdownTypes = @('markdown')
$labelledTypes = @('textarea', 'input', 'dropdown', 'checkboxes')

$fail = 0
$pass = 0
function Check {
    param([string]$N, [bool]$Ok, [string]$D = '')
    if ($Ok) { $script:pass++; if (-not $Quiet) { Write-Host ("  PASS  {0,-62} {1}" -f $N, $D) -ForegroundColor Green } }
    else { $script:fail++; Write-Host ("  FAIL  {0,-62} {1}" -f $N, $D) -ForegroundColor Red }
}

$yamlFiles = @(Get-ChildItem -LiteralPath (Join-Path $root '.github') -Recurse -Include '*.yml', '*.yaml' -File -ErrorAction SilentlyContinue | Sort-Object FullName)
if ($yamlFiles.Count -eq 0) {
    if (-not $Quiet) { Write-Host 'No .github YAML files found.' -ForegroundColor Yellow }
    exit 0
}

foreach ($f in $yamlFiles) {
    $rel = $f.FullName.Substring($root.Length).TrimStart([IO.Path]::DirectorySeparatorChar, '/')
    $lines = @(Get-Content -LiteralPath $f.FullName)

    # ---- 1. escape safety, block-scalar aware ------------------------------
    # Track whether we are inside a block scalar (| or >). While inside, content
    # is literal and backslashes are NOT escapes.
    $badEscapes = @()
    $inBlock = $false
    $blockIndent = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $ln = $lines[$i]
        $indent = ([regex]::Match($ln, '^\s*')).Length

        if ($inBlock) {
            # A block scalar ends at the first non-blank line that dedents to or
            # past the parent key's indentation.
            if ($ln.Trim() -ne '' -and $indent -le $blockIndent) { $inBlock = $false }
            else { continue }
        }

        # Detect the start of a block scalar: `key: |`, `key: |-`, `key: >`, etc.
        if ($ln -match '^\s*[^#]*?:\s*[|>][+-]?\s*(#.*)?$') {
            $inBlock = $true
            $blockIndent = $indent
            continue
        }

        # Outside a block scalar: validate escapes inside double-quoted scalars.
        if ($ln.TrimStart().StartsWith('#')) { continue }
        foreach ($m in [regex]::Matches($ln, '"((?:[^"\\]|\\.)*)"')) {
            $body = $m.Groups[1].Value
            foreach ($em in [regex]::Matches($body, '\\(.)')) {
                if ($validEscapes -notcontains $em.Groups[1].Value) {
                    $badEscapes += ("line {0}: invalid escape '\{1}'" -f ($i + 1), $em.Groups[1].Value)
                }
            }
        }
    }
    Check "$rel : no invalid escapes in double-quoted scalars" ($badEscapes.Count -eq 0) `
        $(if ($badEscapes) { ($badEscapes | Select-Object -First 3) -join '; ' } else { '' })

    # ---- 2. issue-form schema ----------------------------------------------
    if ($rel -notmatch 'ISSUE_TEMPLATE') { continue }
    if ($f.Name -eq 'config.yml') {
        # config.yml is a different schema: blank_issues_enabled / contact_links.
        $unknown = @()
        foreach ($ln in $lines) {
            if ($ln -match '^([A-Za-z_][A-Za-z0-9_]*):') {
                $k = $Matches[1]
                if ($k -notin @('blank_issues_enabled', 'contact_links')) { $unknown += $k }
            }
        }
        Check "$rel : only known top-level keys" ($unknown.Count -eq 0) $(if ($unknown) { $unknown -join ', ' } else { '' })

        $needsName = 0; $hasName = 0; $needsUrl = 0; $hasUrl = 0; $needsAbout = 0; $hasAbout = 0
        foreach ($ln in $lines) {
            if ($ln -match '^\s*-\s+name:') { $needsName++; $hasName++ }
            elseif ($ln -match '^\s*name:') { $hasName++ }
            if ($ln -match '^\s*url:') { $hasUrl++ }
            if ($ln -match '^\s*about:') { $hasAbout++ }
        }
        $links = @($lines | Where-Object { $_ -match '^\s*-\s+name:' }).Count
        Check "$rel : every contact link has name/url/about" (($hasUrl -ge $links) -and ($hasAbout -ge $links) -and ($links -gt 0)) `
            ("$links link(s), url=$hasUrl, about=$hasAbout")
        continue
    }

    # A form template: name, description, body[] with typed elements.
    $hasName = ($lines | Where-Object { $_ -match '^name:' }).Count -gt 0
    $hasDesc = ($lines | Where-Object { $_ -match '^description:' }).Count -gt 0
    $hasBody = ($lines | Where-Object { $_ -match '^body:' }).Count -gt 0
    Check "$rel : has required name/description/body" ($hasName -and $hasDesc -and $hasBody)

    # Walk the body elements. Each starts with `  - type: <t>`; its attributes
    # follow at deeper indentation.
    $problems = @()
    $count = 0
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $m = [regex]::Match($lines[$i], '^\s*-\s+type:\s*([A-Za-z]+)\s*$')
        if (-not $m.Success) { continue }
        $count++
        $type = $m.Groups[1].Value
        $baseIndent = ([regex]::Match($lines[$i], '^\s*')).Length

        # Collect this element's lines until the next element or a dedent.
        $chunk = New-Object System.Collections.Generic.List[string]
        for ($j = $i + 1; $j -lt $lines.Count; $j++) {
            $lj = $lines[$j]
            if ($lj.Trim() -eq '') { $chunk.Add($lj); continue }
            $ind = ([regex]::Match($lj, '^\s*')).Length
            if ($ind -le $baseIndent) { break }
            $chunk.Add($lj)
        }
        $text = ($chunk -join "`n")

        if ($markdownTypes -contains $type) {
            if ($text -notmatch '(?m)^\s*value:') { $problems += ("element {0} (markdown) has no attributes.value" -f $count) }
            if ($text -match '(?m)^\s*content:') { $problems += ("element {0} (markdown) uses 'content' - GitHub expects 'value'" -f $count) }
        }
        elseif ($labelledTypes -contains $type) {
            if ($text -notmatch '(?m)^\s*label:') { $problems += ("element {0} ($type) has no attributes.label" -f $count) }
            if ($type -eq 'dropdown' -and $text -notmatch '(?m)^\s*options:') {
                $problems += ("element {0} (dropdown) has no options" -f $count)
            }
            if ($type -eq 'checkboxes' -and $text -notmatch '(?m)^\s*options:') {
                $problems += ("element {0} (checkboxes) has no options" -f $count)
            }
        }
        else {
            $problems += ("element {0} has unknown type '{1}'" -f $count, $type)
        }
        $i = $i  # continue scanning; nested elements are not a thing here
    }
    Check "$rel : $count body element(s) match the issue-form schema" ($problems.Count -eq 0) `
        $(if ($problems) { ($problems | Select-Object -First 3) -join '; ' } else { '' })
}

if (-not $Quiet) {
    Write-Host ''
    Write-Host ('=' * 78) -ForegroundColor Cyan
    Write-Host ("  TEMPLATE/YAML LINT: {0} file(s), {1} passed, {2} failed" -f $yamlFiles.Count, $pass, $fail) -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
    Write-Host ('=' * 78) -ForegroundColor Cyan
}
exit $(if ($fail -eq 0) { 0 } else { 1 })

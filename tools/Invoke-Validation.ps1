#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Validation gate: parse every PowerShell source file and report syntax errors.

.DESCRIPTION
    Part of the build quality gate. Parses each file with the PowerShell AST
    parser (no execution, so it is safe on any platform) and, with -Deep, also
    assembles the distributables and parses those too.

    Exit code 0 = everything parses. 1 = at least one syntax error.

.PARAMETER Root
    Repository root. Defaults to the parent of the directory holding this script.

.PARAMETER Deep
    Also run tools/Build.ps1 first and parse the generated distributables.

.EXAMPLE
    ./tools/Invoke-Validation.ps1
    ./tools/Invoke-Validation.ps1 -Deep
#>
[CmdletBinding()]
param(
    [string]$Root,
    [switch]$Deep,
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($Root)) {
    $Root = Split-Path -Parent $PSScriptRoot
}
$Root = (Resolve-Path -LiteralPath $Root).Path

function Write-Step { param([string]$Text) if (-not $Quiet) { Write-Host $Text -ForegroundColor Cyan } }
function Write-Ok   { param([string]$Text) if (-not $Quiet) { Write-Host "  OK   $Text" -ForegroundColor Green } }
function Write-Bad  { param([string]$Text) Write-Host "  FAIL $Text" -ForegroundColor Red }

if ($Deep) {
    Write-Step '[1/2] Building distributables...'
    & (Join-Path $Root 'tools/Build.ps1') -Root $Root -SkipTests | Out-Null
}

Write-Step "Parsing PowerShell sources under $Root ..."
$targets = New-Object System.Collections.Generic.List[string]
foreach ($dir in @('src', 'tools', 'tests')) {
    $p = Join-Path $Root $dir
    if (Test-Path -LiteralPath $p) {
        foreach ($f in (Get-ChildItem -LiteralPath $p -Recurse -Filter *.ps1 -File -ErrorAction SilentlyContinue)) {
            $targets.Add($f.FullName)
        }
    }
}
foreach ($extra in @('Clear-PrintSpoolerQueue.ps1')) {
    $p = Join-Path (Join-Path $Root 'dist') $extra
    if (Test-Path -LiteralPath $p) { $targets.Add($p) }
}
$moduleDir = Join-Path (Join-Path $Root 'dist') 'PrintSpoolerCleanup'
if (Test-Path -LiteralPath $moduleDir) {
    foreach ($f in (Get-ChildItem -LiteralPath $moduleDir -Recurse -Include *.ps1, *.psm1, *.psd1 -File -ErrorAction SilentlyContinue)) {
        $targets.Add($f.FullName)
    }
}

$failed = 0
$totalErrors = 0
foreach ($file in ($targets | Sort-Object -Unique)) {
    $tokens = $null; $errors = $null
    try {
        [void][System.Management.Automation.Language.Parser]::ParseFile($file, [ref]$tokens, [ref]$errors)
    }
    catch {
        Write-Bad "$file (exception: $($_.Exception.Message))"
        $failed++; continue
    }
    if ($errors -and $errors.Count -gt 0) {
        $failed++
        $totalErrors += $errors.Count
        Write-Bad ($file.Substring($Root.Length).TrimStart([IO.Path]::DirectorySeparatorChar, '/'))
        foreach ($e in ($errors | Select-Object -First 5)) {
            Write-Host ("         line {0}, col {1}: {2}" -f $e.Extent.StartLineNumber, $e.Extent.StartColumnNumber, $e.Message) -ForegroundColor DarkRed
        }
        if ($errors.Count -gt 5) { Write-Host ("         ... and {0} more" -f ($errors.Count - 5)) -ForegroundColor DarkRed }
    }
    else {
        Write-Ok ($file.Substring($Root.Length).TrimStart([IO.Path]::DirectorySeparatorChar, '/'))
    }
}

Write-Host ''
Write-Host ("Parsed {0} file(s); {1} failed with {2} error(s)." -f $targets.Count, $failed, $totalErrors) -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Red' })
if ($failed -gt 0) { exit 1 }
exit 0

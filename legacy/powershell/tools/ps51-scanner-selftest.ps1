$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$scan = Join-Path $root 'tools/Invoke-PS51CompatibilityScan.ps1'
$exe = (Get-Process -Id $PID).Path

function Invoke-Scan {
    param([string]$Dir)
    $out = Join-Path ([IO.Path]::GetTempPath()) ("psq51-{0}.out" -f [Guid]::NewGuid().ToString('N'))
    $proc = Start-Process -FilePath $exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scan, '-SourceRoots', $Dir) `
        -NoNewWindow -Wait -PassThru -RedirectStandardOutput $out
    $text = if (Test-Path -LiteralPath $out) { Get-Content -LiteralPath $out -Raw } else { '' }
    Remove-Item -LiteralPath $out -Force -ErrorAction SilentlyContinue
    return [pscustomobject]@{ ExitCode = $proc.ExitCode; Output = $text }
}

$fail = 0
$pass = 0
function Expect-Fires {
    param([string]$Label, [string]$Code, [string]$MustMention)
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("psq51-bad-{0}" -f [Guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $tmp -Force
    try {
        Set-Content -LiteralPath (Join-Path $tmp 'Bad.ps1') -Value $Code -Encoding UTF8
        $r = Invoke-Scan -Dir $tmp
        $ok = ($r.ExitCode -eq 1) -and ($r.Output -match [regex]::Escape($MustMention))
        if ($ok) { $script:pass++; Write-Host ("  FIRES   {0}" -f $Label) -ForegroundColor Green }
        else {
            $script:fail++
            Write-Host ("  MISS    {0}" -f $Label) -ForegroundColor Red
            Write-Host ("          expected exit 1 and text '{0}'; got exit {1}" -f $MustMention, $r.ExitCode) -ForegroundColor DarkRed
            if ($r.Output) { Write-Host ("          output: {0}" -f ($r.Output -replace '\s+', ' ').Substring(0, [Math]::Min(200, $r.Output.Length))) -ForegroundColor DarkYellow }
        }
    }
    finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

function Expect-Quiet {
    param([string]$Label, [string]$Code)
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("psq51-ok-{0}" -f [Guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $tmp -Force
    try {
        Set-Content -LiteralPath (Join-Path $tmp 'Ok.ps1') -Value $Code -Encoding UTF8
        $r = Invoke-Scan -Dir $tmp
        if ($r.ExitCode -eq 0) { $script:pass++; Write-Host ("  QUIET   {0}" -f $Label) -ForegroundColor Green }
        else {
            $script:fail++
            Write-Host ("  FALSE+  {0}" -f $Label) -ForegroundColor Red
            Write-Host ("          {0}" -f (($r.Output -split "`n" | Where-Object { $_ -match '\.ps1:' }) -join ' | ')) -ForegroundColor DarkYellow
        }
    }
    finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

Write-Host '=== the scanner must FIRE on each PS7-only class ===' -ForegroundColor Cyan

Expect-Fires -Label "null-coalescing ??" -MustMention "null-coalescing operator" -Code @'
function Test-A { param($x) return $x ?? 'fallback' }
'@

Expect-Fires -Label "null-coalescing assignment ??=" -MustMention "null-coalescing assignment" -Code @'
function Test-B { param($x) $x ??= 5; return $x }
'@

Expect-Fires -Label "ternary ? :" -MustMention "ternary operator" -Code @'
function Test-C { param($x) return ($x -gt 1) ? 'big' : 'small' }
'@

Expect-Fires -Label "null-conditional member ?." -MustMention "null-conditional access" -Code @'
function Test-D { param($obj) return $obj?.Name }
'@

Expect-Fires -Label '$IsWindows automatic variable' -MustMention "does not exist on 5.1" -Code @'
function Test-E { if ($IsWindows) { return 'win' } return 'other' }
'@

Expect-Fires -Label '$IsLinux automatic variable' -MustMention "does not exist on 5.1" -Code @'
function Test-F { return $IsLinux }
'@

Expect-Fires -Label 'UNGUARDED ForEach-Object -Parallel' -MustMention "UNGUARDED" -Code @'
function Test-G { param($items) return $items | ForEach-Object -ThrottleLimit 4 -Parallel { $_ } }
'@

Expect-Fires -Label 'clean{} block (7.4+)' -MustMention "clean{}" -Code @'
function Test-H { param($x) begin { } process { } end { } clean { } }
'@

Expect-Fires -Label 'ConvertFrom-Json -AsHashtable' -MustMention "AsHashtable" -Code @'
function Test-I { param($j) return ConvertFrom-Json -InputObject $j -AsHashtable }
'@

Expect-Fires -Label 'Join-Path -AdditionalChildPath' -MustMention "AdditionalChildPath" -Code @'
function Test-J { return Join-Path -Path C:\ -ChildPath a -AdditionalChildPath b }
'@

Write-Host ''
Write-Host '=== the scanner must stay QUIET on legitimate 5.1 code ===' -ForegroundColor Cyan

Expect-Quiet -Label 'GUARDED -Parallel with a 5.1 fallback (the real pattern)' -Code @'
function Test-Ok1 {
    param($items)
    if ($items.Count -gt 1 -and (Test-PowerShellSevenFeatures)) {
        return $items | ForEach-Object -ThrottleLimit 4 -Parallel { $_ }
    }
    return $items | ForEach-Object { $_ }
}
'@

Expect-Quiet -Label 'GUARDED -Parallel via a $PSVersionTable condition' -Code @'
function Test-Ok2 {
    param($items)
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        return $items | ForEach-Object -ThrottleLimit 2 -Parallel { $_ }
    }
    return $items
}
'@

Expect-Quiet -Label 'IsWindows as a hashtable KEY (not a variable)' -Code @'
function Test-Ok3 { return [pscustomobject]@{ IsWindows = (Test-RunningOnWindows); IsLinux = $false } }
'@

Expect-Quiet -Label "IsWindows as a STRING passed to Get-Variable (the safe probe)" -Code @'
function Test-Ok4 {
    $v = Get-Variable -Name 'IsWindows' -Scope Global -ErrorAction SilentlyContinue
    return ($null -ne $v)
}
'@

Expect-Quiet -Label '-Parallel mentioned in help text and strings only' -Code @'
function Test-Ok5 {
    <#
    .SYNOPSIS
        Uses ForEach-Object -Parallel on PowerShell 7 and ?? operators elsewhere.
    #>
    param($x)
    Write-Host 'ForEach-Object -Parallel requires PS7; the ?? operator too.'
    return $x
}
'@

Expect-Quiet -Label 'plain 5.1-era code' -Code @'
function Test-Ok6 {
    param([string]$Path)
    $items = @()
    foreach ($f in @(Get-ChildItem -LiteralPath $Path -File)) { $items += $f.Name }
    $null = New-Item -ItemType Directory -Path $Path -Force
    return ($items -join ',')
}
'@

Write-Host ''
Write-Host ('=' * 78) -ForegroundColor Cyan
Write-Host ("  PS5.1 SCANNER SELF-TEST: {0} passed, {1} failed" -f $pass, $fail) -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
Write-Host ('=' * 78) -ForegroundColor Cyan
exit $(if ($fail -eq 0) { 0 } else { 1 })

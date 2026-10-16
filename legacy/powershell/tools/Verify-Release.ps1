$root = Split-Path -Parent $PSScriptRoot
$dist = Join-Path $root 'dist'
$fail = 0
function Check { param([string]$N, [bool]$Ok, [string]$D = '')
    if ($Ok) { Write-Host ("  PASS  {0,-56} {1}" -f $N, $D) -ForegroundColor Green }
    else { Write-Host ("  FAIL  {0,-56} {1}" -f $N, $D) -ForegroundColor Red; $script:fail++ }
}

Write-Host '=== Release artifact integrity ===' -ForegroundColor Cyan

# 1. Every artifact listed in checksums.txt must exist and match.
$sums = @{}
foreach ($line in @(Get-Content -LiteralPath (Join-Path $dist 'checksums.txt'))) {
    if ($line -match '^([0-9a-fA-F]{64})\s+(.+)$') { $sums[$Matches[2].Trim()] = $Matches[1] }
}
Check 'checksums.txt parsed' ($sums.Count -ge 4) "$($sums.Count) entries"
foreach ($name in @($sums.Keys | Sort-Object)) {
    $path = Join-Path $dist $name
    if (-not (Test-Path -LiteralPath $path)) { Check "artifact present: $name" $false 'missing'; continue }
    $actual = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
    Check "SHA-256 matches: $name" ($actual -eq $sums[$name].ToLowerInvariant())
}

# 2. The ZIP must contain the module, the script, the launcher and the config.
# A verification tool that THROWS instead of reporting is itself a defect: the
# operator gets a MethodInvocationException with no indication of what was
# actually wrong. Report the missing archive as a failed check and skip only the
# ZIP-specific assertions.
$zip = @(Get-ChildItem -LiteralPath $dist -Filter 'PrintSpoolerCleanup-*.zip' -ErrorAction SilentlyContinue)[0]
if ($null -eq $zip) {
    Check 'Release ZIP is present' $false 'no PrintSpoolerCleanup-*.zip in dist/ (built with -NoZip?)'
    $archive = $null
}
else {
    Check 'Release ZIP is present' $true $zip.Name
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = if ($zip) { [System.IO.Compression.ZipFile]::OpenRead($zip.FullName) } else { $null }
if ($archive) {
try {
    $names = @($archive.Entries | ForEach-Object { $_.FullName })
    Check 'ZIP contains the standalone script' ($names -contains 'Clear-PrintSpoolerQueue.ps1')
    Check 'ZIP contains the launcher' ($names -contains 'Clear-PrintSpoolerQueue.bat')
    Check 'ZIP contains the module manifest' ($names -contains 'PrintSpoolerCleanup/PrintSpoolerCleanup.psd1')
    Check 'ZIP contains the module script' ($names -contains 'PrintSpoolerCleanup/PrintSpoolerCleanup.psm1')
    Check 'ZIP contains the config template' ($names -contains 'PrintSpoolerCleanup.config.psd1')
    Check 'ZIP contains README and LICENSE' (($names -contains 'README.md') -and ($names -contains 'LICENSE'))
    Check 'ZIP contains no log/runtime junk' (@($names | Where-Object { $_ -match 'Logs/|\.update/|\.log$|update-state' }).Count -eq 0) ($names -join ', ')
    # No zip-slip: every entry name must be a plain relative path.
    $traversal = @($names | Where-Object { $_.StartsWith('..') -or $_.StartsWith('/') -or $_.StartsWith('\') -or $_.Contains(':\') -or $_.Contains('/../') })
    Check 'ZIP has no path-traversal entries' ($traversal.Count -eq 0) ($traversal -join ', ')
}
finally { $archive.Dispose() }
}

# 3. Version consistency across every shipped surface.
$ver = ([regex]::Match((Get-Content -LiteralPath (Join-Path $dist 'Clear-PrintSpoolerQueue.ps1') -Raw), "\`$Script:ScriptVersion\s+=\s+'([^']+)'")).Groups[1].Value
Check 'Script carries a version stamp' (-not [string]::IsNullOrWhiteSpace($ver)) $ver
$bat = Get-Content -LiteralPath (Join-Path $dist 'Clear-PrintSpoolerQueue.bat') -Raw
Check 'Launcher version matches the script' ($bat -match [regex]::Escape("SCRIPT_VERSION=$ver"))
$psd = Get-Content -LiteralPath (Join-Path $dist 'PrintSpoolerCleanup/PrintSpoolerCleanup.psd1') -Raw
Check 'Manifest version matches the script' ($psd -match [regex]::Escape("'$ver'"))
$psm = Get-Content -LiteralPath (Join-Path $dist 'PrintSpoolerCleanup/PrintSpoolerCleanup.psm1') -Raw
Check 'Module script version matches' ($psm -match [regex]::Escape("ScriptVersion       = '$ver'"))
if ($zip) {
    $zipName = $zip.Name
    Check 'ZIP filename carries the version' ($zipName -match [regex]::Escape($ver)) $zipName
}
$cfg = Get-Content -LiteralPath (Join-Path $dist 'PrintSpoolerCleanup.config.psd1') -Raw
Check 'Config template references the version' ($cfg -match [regex]::Escape($ver))

# 4. No build token survived anywhere.
foreach ($f in @('Clear-PrintSpoolerQueue.ps1', 'Clear-PrintSpoolerQueue.bat', 'PrintSpoolerCleanup/PrintSpoolerCleanup.psd1', 'PrintSpoolerCleanup/PrintSpoolerCleanup.psm1', 'PrintSpoolerCleanup.config.psd1')) {
    $t = Get-Content -LiteralPath (Join-Path $dist $f) -Raw
    # Only BUILD-time tokens are a defect. @@SCRIPT_NAME@@, @@SCRIPT_VERSION@@ and
    # the other payload tokens are substituted at RUNTIME by
    # Resolve-PayloadScriptBlock and must survive into the .ps1/.psm1 - the build
    # stage asserts that exact whitelist separately.
    $left = @([regex]::Matches($t, '@@BUILD_STAMP@@') | ForEach-Object { $_.Value } | Select-Object -Unique)
    Check "No leftover build tokens: $f" ($left.Count -eq 0) ($left -join ',')
}

# 5. The dist tree contains artifacts only.
$junk = @(Get-ChildItem -LiteralPath $dist -Recurse -Force | Where-Object { $_.Name -in @('Logs', '.update', 'update-state.json', 'pending-update.json') })
Check 'dist/ contains no runtime junk' ($junk.Count -eq 0) (($junk | ForEach-Object { $_.Name }) -join ', ')

# 6. Documentation set is complete.
$wiki = @(Get-ChildItem -LiteralPath (Join-Path $root 'docs/wiki') -Filter '*.md')
Check 'Wiki has at least 19 content pages' ($wiki.Count -ge 20) "$($wiki.Count) files"
Check 'Wiki has Home and _Sidebar' (($wiki.Name -contains 'Home.md') -and ($wiki.Name -contains '_Sidebar.md'))
foreach ($doc in @('README.md', 'CHANGELOG.md', 'SECURITY.md', 'CONTRIBUTING.md', 'CODE_OF_CONDUCT.md', 'LICENSE', '.editorconfig', '.gitignore')) {
    Check "Governance file present: $doc" (Test-Path -LiteralPath (Join-Path $root $doc))
}
foreach ($doc in @('Exit-Codes.md', 'Report-Schema.md', 'REFACTOR-PLAN.md', 'HANDOVER.md')) {
    Check "Engineering doc present: docs/$doc" (Test-Path -LiteralPath (Join-Path $root "docs/$doc"))
}
Check 'CI workflow present' (Test-Path -LiteralPath (Join-Path $root '.github/workflows/ci.yml'))
Check 'Issue templates present' (@(Get-ChildItem -LiteralPath (Join-Path $root '.github/ISSUE_TEMPLATE') -Filter '*.yml').Count -ge 3)

# 7. License consistency (finding A15: the README claimed Public Domain).
$lic = Get-Content -LiteralPath (Join-Path $root 'LICENSE') -Raw
$readme = Get-Content -LiteralPath (Join-Path $root 'README.md') -Raw
Check 'LICENSE is GPL v3' ($lic -match 'GNU GENERAL PUBLIC LICENSE')
# The README legitimately mentions "Public Domain" once, in the note explaining
# that the old badge was wrong. What must not exist is a Public Domain BADGE or
# licence claim.
$badgeLine = @($readme -split "\r?\n" | Where-Object { $_ -match '^\[!' }) -join ' '
Check 'No Public Domain badge in the README' ($badgeLine -notmatch 'Public%20Domain|Public Domain')
Check 'No Public Domain licence claim outside the correction note' (([regex]::Matches($readme, 'Public Domain')).Count -le 1)
Check 'README states GPL-3.0' ($readme -match 'GPL-3\.0')

Write-Host ''
Write-Host ('=' * 78) -ForegroundColor Cyan
Write-Host ("  ARTIFACT INTEGRITY: {0} failure(s)" -f $fail) -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
Write-Host ('=' * 78) -ForegroundColor Cyan
exit $(if ($fail -eq 0) { 0 } else { 1 })

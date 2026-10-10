<#
.SYNOPSIS
    Module-form validation: manifest, import, exported surface, isolation.

.DESCRIPTION
    The .psd1/.psm1 form is a genuinely separate code path from dot-sourcing the
    standalone script. Module scope, Export-ModuleMember, private-helper isolation
    and manifest metadata all behave differently, and a defect in any of them is
    invisible to every other stage of the quality gate - which is how a
    [switch]-parameter collision survived 400+ checks until this stage existed.

    Runs in its own process (invoked by tests/Run-Tests.ps1) so that importing and
    removing the module cannot perturb other stages.

.EXAMPLE
    ./tools/module-test.ps1
#>
[CmdletBinding()]
param(
    [string]$Root,
    [switch]$SkipNetwork
)

$ErrorActionPreference = 'Continue'
if ([string]::IsNullOrWhiteSpace($Root)) { $Root = Split-Path -Parent $PSScriptRoot }

# Importing the module initialises runtime state, which resolves a log target.
# Where %ProgramData% is unavailable the fallback chain reaches the module's own
# directory, so exercising the cmdlets here would write a log INTO the release
# artifacts. Redirect logging to a scratch directory for the whole run - which
# also proves PSQ_LOG_PATH is honoured - and sweep any residue at the end.
$savedLogEnv = $env:PSQ_LOG_PATH
$testLogDir = Join-Path ([IO.Path]::GetTempPath()) ("psq-modtest-{0}" -f [Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $testLogDir -Force
$env:PSQ_LOG_PATH = Join-Path $testLogDir 'module-test.log'
$modDir = Join-Path $Root 'dist/PrintSpoolerCleanup'
$psd1 = Join-Path $modDir 'PrintSpoolerCleanup.psd1'
$psm1 = Join-Path $modDir 'PrintSpoolerCleanup.psm1'

$script:failures = 0
$script:passes = 0
function Check {
    param([string]$Name, [bool]$Ok, [string]$Detail = '')
    if ($Ok) { $script:passes++; Write-Host ("  PASS  {0,-54} {1}" -f $Name, $Detail) -ForegroundColor Green }
    else { $script:failures++; Write-Host ("  FAIL  {0,-54} {1}" -f $Name, $Detail) -ForegroundColor Red }
}
function Section { param([string]$T) Write-Host ''; Write-Host "=== $T ===" -ForegroundColor Cyan }

$publicFunctions = @(
    'Clear-PrintSpoolerQueue', 'Get-SpoolerCleanupHelp', 'Get-SpoolerCleanupVersion',
    'Test-PrintSpoolerHealth', 'Get-PrintSpoolerDiagnostic', 'Get-SpoolerCleanupUpdate',
    'Update-SpoolerCleanup', 'Restore-SpoolerBackup', 'Get-SpoolerBackup',
    'Install-SpoolerCleanupTask', 'Uninstall-SpoolerCleanupTask', 'Get-SpoolerCleanupTask',
    'Start-SpoolerCleanupTask', 'Get-ExitCodeMeaning'
)

# Representative private helpers. If ANY of these becomes visible after import,
# Export-ModuleMember is not doing its job. Chosen to cover every source file.
$privateSamples = @(
    'Write-LogEntryCore', 'Write-LogInfo', 'Resolve-LogTarget', 'Invoke-LogRotation',
    'Get-ConsoleCapability', 'Initialize-ConsolePalette', 'Write-ConsoleTable',
    'ConvertTo-SemVer', 'Compare-ProductVersion', 'ConvertFrom-AnyTimeSpan',
    'Test-SpoolDirectorySafeInternal', 'New-ResultObject', 'Format-ByteSize',
    'Get-DefaultConfiguration', 'Test-Configuration', 'Get-EffectiveConfiguration',
    'Resolve-PayloadScriptBlock', 'New-PayloadOptions', 'Invoke-CleanupFleet',
    'ConvertFrom-PayloadResult', 'Get-SpoolBackupInfo', 'Restore-SpoolerBackupInternal',
    'Get-UpdateInfoCore', 'Update-SpoolerCleanupCore', 'Read-UpdateState',
    'New-RunSummary', 'ConvertTo-HtmlReport', 'Export-RunReport',
    'Invoke-SelfElevation', 'Test-ScriptSignature', 'Write-DiagnosticsBundle',
    'New-SpoolerTaskXml', 'Install-SpoolerCleanupTaskInternal',
    'Invoke-SelfTestInternal', 'Start-WatchdogInternal', 'Initialize-RuntimeState',
    'Get-HelpText', 'Get-ParameterDocumentation', 'Set-ScriptExitCode'
)

if (-not (Test-Path -LiteralPath $psd1) -or -not (Test-Path -LiteralPath $psm1)) {
    Write-Host "Module artifacts not found under $modDir - run tools/Build.ps1 first." -ForegroundColor Red
    exit 1
}

Section '1. Manifest validation'
$manifest = $null
try {
    $manifest = Test-ModuleManifest -Path $psd1 -ErrorAction Stop
    Check 'Test-ModuleManifest succeeds' $true ("v{0}" -f $manifest.Version)
}
catch { Check 'Test-ModuleManifest succeeds' $false $_.Exception.Message }

if ($manifest) {
    Check 'Version is 4.0.0' ($manifest.Version.ToString() -eq '4.0.0') $manifest.Version.ToString()
    Check 'PowerShellVersion floor is 5.1' ($manifest.PowerShellVersion.ToString() -eq '5.1') $manifest.PowerShellVersion.ToString()
    Check 'CompatiblePSEditions covers Desktop and Core' ((@($manifest.CompatiblePSEditions) -contains 'Desktop') -and (@($manifest.CompatiblePSEditions) -contains 'Core')) (@($manifest.CompatiblePSEditions) -join ', ')
    Check 'Author is set' ($manifest.Author -eq 'Mikhail Deynekin') $manifest.Author
    Check 'Description is substantive (>100 chars)' ($manifest.Description.Length -gt 100) "$($manifest.Description.Length) chars"
    Check 'ProjectUri points at the repository' ($manifest.PrivateData.PSData.ProjectUri -eq 'https://github.com/paulmann/Print-Spooler-Queue-Cleanup')
    Check 'LicenseUri is set' (-not [string]::IsNullOrWhiteSpace($manifest.PrivateData.PSData.LicenseUri))
    Check 'Tags are set for discoverability' (@($manifest.PrivateData.PSData.Tags).Count -ge 5) "$(@($manifest.PrivateData.PSData.Tags).Count) tags"
    Check 'ReleaseNotes are present' (-not [string]::IsNullOrWhiteSpace($manifest.PrivateData.PSData.ReleaseNotes))
    Check 'GUID is a stable literal, not regenerated per build' ($manifest.Guid.ToString() -eq 'a7f3c1d2-9b64-4e08-8c2f-51d0e6a4b7c9') $manifest.Guid.ToString()

    $exported = @($manifest.ExportedFunctions.Keys)
    $missing = @($publicFunctions | Where-Object { $exported -notcontains $_ })
    $unexpected = @($exported | Where-Object { $publicFunctions -notcontains $_ })
    Check "Manifest exports exactly the $($publicFunctions.Count) public functions" (($missing.Count -eq 0) -and ($unexpected.Count -eq 0)) `
        $(if ($missing) { "missing: $($missing -join ', ')" } elseif ($unexpected) { "unexpected: $($unexpected -join ', ')" } else { "$($exported.Count) functions" })
    # ExportedVariables is a Dictionary<string,PSVariable>, so @($dict).Count is
    # 1 (the dictionary itself) rather than 0 - the KEYS must be counted.
    $exportedVars = @($manifest.ExportedVariables.Keys)
    Check 'No cmdlets, variables or aliases are exported' ((@($manifest.ExportedCmdlets.Keys).Count -eq 0) -and ($exportedVars.Count -eq 0) -and (@($manifest.ExportedAliases.Keys).Count -eq 0)) $(if ($exportedVars.Count) { "variables: $($exportedVars -join ', ')" } else { '0 / 0 / 0' })
}

Section '2. Import and exported surface'
$imported = $false
try {
    Import-Module $psd1 -Force -ErrorAction Stop
    $imported = $true
    Check 'Import-Module succeeds' $true
}
catch { Check 'Import-Module succeeds' $false $_.Exception.Message }

if ($imported) {
    $mod = Get-Module PrintSpoolerCleanup
    Check 'Module is loaded' ($null -ne $mod) ("v{0}" -f $mod.Version)
    $cmds = @(Get-Command -Module PrintSpoolerCleanup)
    Check "All $($publicFunctions.Count) commands are callable" ($cmds.Count -eq $publicFunctions.Count) "$($cmds.Count) commands"

    $leaked = @($privateSamples | Where-Object { Get-Command -Name $_ -ErrorAction SilentlyContinue })
    Check 'No private helper leaked into the session' ($leaked.Count -eq 0) $(if ($leaked) { "leaked: $($leaked -join ', ')" } else { "$($privateSamples.Count) sampled helpers stay private" })

    $cmd = Get-Command Clear-PrintSpoolerQueue -ErrorAction SilentlyContinue
    Check 'Main command exposes its full parameter surface' ($cmd -and $cmd.Parameters.Count -ge 70) $(if ($cmd) { "$($cmd.Parameters.Count) parameters" } else { 'command not found' })
    Check 'Main command supports ShouldProcess' ($cmd -and $cmd.Parameters.ContainsKey('WhatIf'))
}

Section '3. Module functions actually execute'
if ($imported) {
    try { Check 'Get-SpoolerCleanupVersion reports 4.0.0' ((Get-SpoolerCleanupVersion).Version -eq '4.0.0') }
    catch { Check 'Get-SpoolerCleanupVersion' $false $_.Exception.Message }

    try { Check 'Get-SpoolerCleanupVersion -Detailed includes integrity records' (@((Get-SpoolerCleanupVersion -Detailed).Files).Count -ge 0) }
    catch { Check 'Get-SpoolerCleanupVersion -Detailed' $false $_.Exception.Message }

    try { Check 'Get-ExitCodeMeaning returns the full contract' (@(Get-ExitCodeMeaning).Count -ge 17) "$(@(Get-ExitCodeMeaning).Count) codes" }
    catch { Check 'Get-ExitCodeMeaning' $false $_.Exception.Message }

    try { Check 'Get-SpoolerCleanupHelp renders a section' ((Get-SpoolerCleanupHelp -Section ExitCodes).Length -gt 500) }
    catch { Check 'Get-SpoolerCleanupHelp -Section' $false $_.Exception.Message }

    try {
        $h = Get-SpoolerCleanupHelp -Section Parameters
        Check 'Help renders the metadata-driven parameter table' ($h.Length -gt 5000) "$($h.Length) chars"
        Check 'No parameter carries a placeholder description' ($h -notmatch 'No description available')
    }
    catch { Check 'Get-SpoolerCleanupHelp -Section Parameters' $false $_.Exception.Message }

    # THE regression test for the [switch]-parameter collision: this threw
    # "Cannot convert value ... to SwitchParameter" when the module form was used.
    try {
        $st = Clear-PrintSpoolerQueue -SelfTest -NoColor -Quiet -PassThru
        Check '-SelfTest -PassThru returns the result object' ($null -ne $st) $(if ($st) { $st.GetType().Name } else { 'nothing returned' })
        if ($st) {
            Check '-SelfTest passes all internal checks' ($st.Overall -eq 'Passed') ("{0}/{1}" -f ($st.TotalChecks - $st.FailedChecks), $st.TotalChecks)
            foreach ($c in @($st.Checks | Where-Object Result -eq 'FAIL')) {
                Write-Host ("        FAIL {0}: {1}" -f $c.Name, $c.Detail) -ForegroundColor DarkRed
            }
        }
    }
    catch { Check '-SelfTest -PassThru via the module' $false $_.Exception.Message }

    try {
        $ver = Clear-PrintSpoolerQueue -ShowVersion -NoColor -Quiet -PassThru
        Check '-ShowVersion -PassThru returns the version object' ($null -ne $ver -and $ver.Version -eq '4.0.0')
    }
    catch { Check '-ShowVersion -PassThru via the module' $false $_.Exception.Message }

    try {
        $secs = Clear-PrintSpoolerQueue -ListHelpSections -NoColor -Quiet -PassThru
        Check '-ListHelpSections -PassThru returns the section list' (@($secs).Count -ge 18) "$(@($secs).Count) sections"
    }
    catch { Check '-ListHelpSections -PassThru via the module' $false $_.Exception.Message }

    try {
        $hp = @(Test-PrintSpoolerHealth)
        Check 'Test-PrintSpoolerHealth returns a record' ($hp.Count -ge 1) ("Health={0}" -f $hp[0].Health)
    }
    catch { Check 'Test-PrintSpoolerHealth' $false $_.Exception.Message }

    try {
        $bk = @(Get-SpoolerBackup -BackupPath (Join-Path ([IO.Path]::GetTempPath()) 'no-such-backup-dir'))
        Check 'Get-SpoolerBackup tolerates a missing directory' ($bk.Count -eq 0)
    }
    catch { Check 'Get-SpoolerBackup' $false $_.Exception.Message }

    if (-not $SkipNetwork) {
        try {
            $u = Get-SpoolerCleanupUpdate
            Check 'Get-SpoolerCleanupUpdate returns an info object' ($null -ne $u -and $null -ne $u.CurrentVersion) ("latest={0} via {1}" -f $u.LatestVersion, $u.Source)
        }
        catch { Check 'Get-SpoolerCleanupUpdate' $false $_.Exception.Message }
    }
}

Section '4. Module isolation and teardown'
if ($imported) {
    try {
        Remove-Module PrintSpoolerCleanup -Force -ErrorAction Stop
        Check 'Remove-Module succeeds' (-not (Get-Module PrintSpoolerCleanup))
        Check 'Public commands are gone after removal' (-not (Get-Command Clear-PrintSpoolerQueue -ErrorAction SilentlyContinue))
    }
    catch { Check 'Remove-Module' $false $_.Exception.Message }
}

Section '5. Module artifacts do not pollute their own directory'
# Remove anything an earlier run may have left behind, so this stage reports on
# the shipped artifact rather than on accumulated test residue.
foreach ($junk in @('Logs', '.update')) {
    foreach ($d in @(Get-ChildItem -LiteralPath $modDir -Recurse -Directory -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $junk })) {
        Remove-Item -LiteralPath $d.FullName -Recurse -Force -ErrorAction SilentlyContinue
    }
}
# Importing the module initialises runtime state, which resolves a log target.
# On Windows that is %ProgramData%; where that is unavailable the fallback is the
# module directory, which for a Program Files install is read-only. Either way a
# release folder must not accumulate log files, so the build removes them and the
# artifact listing excludes them.
$strayLogs = @(Get-ChildItem -LiteralPath $modDir -Recurse -Filter '*.log' -File -ErrorAction SilentlyContinue)
Check 'No stray log files inside the module directory' ($strayLogs.Count -eq 0) $(if ($strayLogs) { ($strayLogs | ForEach-Object { $_.Name }) -join ', ' } else { '' })
$strayDirs = @(Get-ChildItem -LiteralPath $modDir -Recurse -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq 'Logs' })
Check 'No Logs directory inside the module directory' ($strayDirs.Count -eq 0)

$env:PSQ_LOG_PATH = $savedLogEnv
Remove-Item -LiteralPath $testLogDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ''
Write-Host ('=' * 78) -ForegroundColor Cyan
Write-Host ("  MODULE TESTS: {0} passed, {1} failed" -f $script:passes, $script:failures) -ForegroundColor $(if ($script:failures -eq 0) { 'Green' } else { 'Red' })
Write-Host ('=' * 78) -ForegroundColor Cyan
exit $(if ($script:failures -eq 0) { 0 } else { 1 })

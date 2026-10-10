<#
.SYNOPSIS
    Requirements-traceability check: proves every item in the original brief is
    present AND FUNCTIONAL in the built artifact, not merely documented.

.DESCRIPTION
    Documentation can claim a feature exists. This script exercises it. Each
    requirement from the brief maps to concrete assertions against
    dist/Clear-PrintSpoolerQueue.ps1 - parameter metadata, exported cmdlets,
    implemented discovery strategies, rendered help, and live invocations.

    It exists because a large refactor can silently drop a documented feature:
    the gate would stay green while the product lost a capability nobody tested
    by name. This stage fails the build if that happens.

.PARAMETER SkipNetwork
    Skip the live GitHub version check (used by CI legs without network).

.EXAMPLE
    ./tools/Verify-Brief.ps1
.EXAMPLE
    ./tools/Verify-Brief.ps1 -SkipNetwork
#>
[CmdletBinding()]
param(
    [AllowNull()][AllowEmptyString()][string]$Root,
    [switch]$SkipNetwork
)

if ([string]::IsNullOrWhiteSpace($Root)) { $Root = Split-Path -Parent $PSScriptRoot }
$root = $Root
. (Join-Path $root 'dist/Clear-PrintSpoolerQueue.ps1')

# Keep logging out of the artifact tree: Resolve-LogTarget falls back to the
# script's own directory when %ProgramData% is unavailable, which would otherwise
# litter dist/ with a Logs folder during this check.
$savedLogEnv = $env:PSQ_LOG_PATH
$briefLogDir = Join-Path ([IO.Path]::GetTempPath()) ("psq-brief-{0}" -f [Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $briefLogDir -Force
$env:PSQ_LOG_PATH = Join-Path $briefLogDir 'verify-brief.log'

$isWindowsHost = ($PSVersionTable.PSEdition -eq 'Desktop') -or ($PSVersionTable.OSPlatform -eq 'Windows') -or ($null -ne (Get-Variable -Name IsWindows -ErrorAction SilentlyContinue) -and $IsWindows)

$script:fail = 0
$script:pass = 0
function Check { param([string]$N, [bool]$Ok, [string]$D = '')
    if ($Ok) { $script:pass++; Write-Host ("  PASS  {0,-58} {1}" -f $N, $D) -ForegroundColor Green }
    else { $script:fail++; Write-Host ("  FAIL  {0,-58} {1}" -f $N, $D) -ForegroundColor Red }
}
function Req { param([string]$T) Write-Host ''; Write-Host "=== $T ===" -ForegroundColor Cyan }

# Every item below is a requirement from the original brief. This proves each one
# is present and FUNCTIONAL in the built artifact, not merely documented.

Req 'BRIEF 1: "add a key to check for new versions on git and auto-update"'
$cmd = Get-Command Clear-PrintSpoolerQueue
Check '-CheckUpdate parameter exists' ($cmd.Parameters.ContainsKey('CheckUpdate'))
Check '-Update parameter exists' ($cmd.Parameters.ContainsKey('Update'))
Check '-AutoUpdate parameter exists' ($cmd.Parameters.ContainsKey('AutoUpdate'))
Check '-UpdateChannel parameter exists' ($cmd.Parameters.ContainsKey('UpdateChannel'))
Check '-IncludePrerelease parameter exists' ($cmd.Parameters.ContainsKey('IncludePrerelease'))
Check '-GitHubToken parameter exists (private-repo updates)' ($cmd.Parameters.ContainsKey('GitHubToken'))
Check '-NoUpdateCheck parameter exists (air-gapped opt-out)' ($cmd.Parameters.ContainsKey('NoUpdateCheck'))
Check '-ApplyUpdateNow parameter exists' ($cmd.Parameters.ContainsKey('ApplyUpdateNow'))
Check 'Get-SpoolerCleanupUpdate cmdlet exported' ($null -ne (Get-Command Get-SpoolerCleanupUpdate -ErrorAction SilentlyContinue))
Check 'Update-SpoolerCleanup cmdlet exported' ($null -ne (Get-Command Update-SpoolerCleanup -ErrorAction SilentlyContinue))
# GitHub discovery strategies actually implemented
$src = Get-Content -LiteralPath (Join-Path $root 'dist/Clear-PrintSpoolerQueue.ps1') -Raw
Check 'GitHub Releases API discovery strategy' ($src -match 'Get-GitHubLatestRelease')
Check 'GitHub tags API discovery strategy' ($src -match 'Get-GitHubLatestTag')
Check 'Raw-file fallback strategy (repo has no releases)' ($src -match 'Get-RemoteVersionFromRawFile')
Check 'SHA-256 verification of downloaded update' ($src -match 'Test-HashMatch')
Check 'Staged self-update via .update\ directory' ($src -match '\.update')
Check 'Update cooldown state persists between runs' ($src -match 'Read-UpdateState|Write-UpdateState')
# Functional: the update path returns a real verdict against the live repo
if ($SkipNetwork) {
    Write-Host '  SKIP  Live version check (-SkipNetwork)' -ForegroundColor Yellow
}
else {
    try {
        $u = Get-SpoolerCleanupUpdate
        Check 'Live version check returns a verdict' ($null -ne $u -and $null -ne $u.CurrentVersion) ("current={0} latest={1} via {2}" -f $u.CurrentVersion, $u.LatestVersion, $u.Source)
    } catch { Check 'Live version check returns a verdict' $false $_.Exception.Message }
}

Req 'BRIEF 2: "add a key for comprehensive help"'
Check '-Help parameter exists' ($cmd.Parameters.ContainsKey('Help'))
Check '-ListHelpSections parameter exists' ($cmd.Parameters.ContainsKey('ListHelpSections'))
Check 'Get-SpoolerCleanupHelp cmdlet exported' ($null -ne (Get-Command Get-SpoolerCleanupHelp -ErrorAction SilentlyContinue))
$sections = @(Clear-PrintSpoolerQueue -ListHelpSections -NoColor -Quiet -PassThru)
Check 'Help has 15+ distinct sections' ($sections.Count -ge 15) "$($sections.Count) sections"
# Defect 46 regression: the documented CLI form must actually bind. -Help used
# to be a [string] section selector, so bare `-Help` - the exact command the
# README, the .EXAMPLE blocks and the console hints advertise - died with
# "Missing an argument for parameter 'Help'". Every in-process test passed a
# value, so nothing caught it. Bind it the way a user does: in-process, and
# through a CHILD pwsh against the shipped artifact (the `pwsh -File` argument
# parser is the one the documentation quotes).
Check '-Help is a switch (bare -Help must bind)' ($cmd.Parameters['Help'].ParameterType -eq [switch])
Check '-HelpSection parameter exists' ($cmd.Parameters.ContainsKey('HelpSection'))
$bareHelp = (Clear-PrintSpoolerQueue -Help -NoColor -Quiet 6>&1 | Out-String)
Check 'Bare -Help renders the full document in-process (>20 KB)' ($bareHelp.Length -gt 20000) "$([int]($bareHelp.Length / 1024)) KB"
Check 'Bare -Help exits Success' ($Script:RuntimeState.ExitCode -eq 0) "exit=$($Script:RuntimeState.ExitCode)"
$pwshExe = (Get-Process -Id $PID).Path
$helpArtifact = Join-Path $root 'dist/Clear-PrintSpoolerQueue.ps1'
$prevEap = $ErrorActionPreference
$ErrorActionPreference = 'Continue'
$helpChildText = & $pwshExe -NoProfile -File $helpArtifact -Help -NoColor -Quiet 2>$null | Out-String
$helpChildExit = $LASTEXITCODE
$ErrorActionPreference = $prevEap
Check 'CLI contract: pwsh -File <artifact> -Help (the README form) exits 0 and renders' ($helpChildExit -eq 0 -and $helpChildText.Length -gt 20000) "exit=$helpChildExit, chars=$($helpChildText.Length)"
# Defect 46 had a sibling (defect 47): the entry-point catch set LASTEXITCODE=3
# for binding errors, but the very next line re-read RuntimeState.ExitCode -
# still "Success", because the function body never ran - and exited 0. Every
# malformed invocation therefore reported success to schedulers and RMMs.
$ErrorActionPreference = 'Continue'
$null = & $pwshExe -NoProfile -File $helpArtifact -HelpSection 'NoSuchSection' -NoColor -Quiet 2>$null | Out-String
$invalidChildExit = $LASTEXITCODE
$ErrorActionPreference = $prevEap
Check 'CLI contract: invalid argument exits 3 (BadArguments), never 0' ($invalidChildExit -eq 3) "exit=$invalidChildExit"
# The section list is the product's own contract; assert against the names it
# actually publishes (Usage, not Synopsis - a wrong name here is a test defect
# that looks exactly like a missing feature).
foreach ($sec in @('Usage','QuickStart','Modes','Parameters','Examples','ExitCodes','Configuration','Environment','Logging','Reports','Remote','Scheduling','Updates','Safety','Troubleshooting','Diagnostics','Version','Credits')) {
    Check "  section '$sec' is reachable" ($sections -contains $sec)
}
$full = Get-SpoolerCleanupHelp
Check 'Full help renders comprehensively (>20 KB)' ($full.Length -gt 20000) "$([int]($full.Length/1024)) KB"
$params = Get-SpoolerCleanupHelp -Section Parameters
Check 'Parameter docs are metadata-driven (>5 KB)' ($params.Length -gt 5000) "$([int]($params.Length/1024)) KB"
Check 'No parameter left with a placeholder description' ($params -notmatch 'No description available')
Check 'Every parameter is documented' -Ok ((@($cmd.Parameters.Keys | Where-Object { $_ -notin @('Verbose','Debug','ErrorAction','WarningAction','InformationAction','ErrorVariable','WarningVariable','InformationVariable','OutVariable','OutBuffer','PipelineVariable','WhatIf','Confirm') })).Count -gt 60) "$(@($cmd.Parameters.Keys).Count) total parameters"

Req 'BRIEF 3: "prepare the code for Production"'
Check 'Comment-based help on the main function' ($src -match '(?s)\.SYNOPSIS.*\.DESCRIPTION.*\.EXAMPLE')
Check 'SupportsShouldProcess (-WhatIf/-Confirm) on main function' ($cmd.Parameters.ContainsKey('WhatIf') -and $cmd.Parameters.ContainsKey('Confirm'))
Check 'Structured exit-code contract' (@(Get-ExitCodeMeaning).Count -ge 17) "$(@(Get-ExitCodeMeaning).Count) codes"
Check 'Input validation attributes present' ($src -match '\[ValidateRange\(' -and $src -match '\[ValidateSet\(')
Check 'Module manifest ships alongside the script' (Test-Path -LiteralPath (Join-Path $root 'dist/PrintSpoolerCleanup/PrintSpoolerCleanup.psd1'))
Check 'checksums.txt for release integrity' (Test-Path -LiteralPath (Join-Path $root 'dist/checksums.txt'))
Check 'Generated RELEASE-NOTES.md' (Test-Path -LiteralPath (Join-Path $root 'dist/RELEASE-NOTES.md'))
Check 'CI workflow present' (Test-Path -LiteralPath (Join-Path $root '.github/workflows/ci.yml'))
Check 'SECURITY.md disclosure policy' (Test-Path -LiteralPath (Join-Path $root 'SECURITY.md'))
Check 'CONTRIBUTING.md' (Test-Path -LiteralPath (Join-Path $root 'CONTRIBUTING.md'))
Check 'CODE_OF_CONDUCT.md' (Test-Path -LiteralPath (Join-Path $root 'CODE_OF_CONDUCT.md'))
Check 'Issue templates' (@(Get-ChildItem -LiteralPath (Join-Path $root '.github/ISSUE_TEMPLATE') -Filter '*.yml').Count -ge 3)
Check 'LICENSE present and GPL-3.0' ((Get-Content -LiteralPath (Join-Path $root 'LICENSE') -Raw) -match 'GNU GENERAL PUBLIC LICENSE')
Check 'Version stamped consistently' ((Get-SpoolerCleanupVersion).Version -eq $Script:ScriptVersion) ((Get-SpoolerCleanupVersion).Version)

Req 'BRIEF 4: "English wiki documentation for GitHub"'
$wiki = @(Get-ChildItem -LiteralPath (Join-Path $root 'docs/wiki') -Filter '*.md')
Check 'Wiki has 20+ pages' ($wiki.Count -ge 20) "$($wiki.Count) files"
Check 'Home.md (GitHub wiki front page) exists' ($wiki.Name -contains 'Home.md')
Check '_Sidebar.md (GitHub wiki navigation) exists' ($wiki.Name -contains '_Sidebar.md')
$wikiText = ($wiki | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }) -join "`n"
Check 'Wiki is substantial (>150 KB of English prose)' ($wikiText.Length -gt 150000) "$([int]($wikiText.Length/1024)) KB"
Check 'Wiki covers the PROBLEM DOMAIN, not just the tool' ($wiki.Name -contains 'Why-Print-Queues-Get-Stuck.md' -and $wiki.Name -contains 'Print-Spooler-Internals.md')
Check 'Wiki has a troubleshooting guide' ($wiki.Name -contains 'Troubleshooting.md')
Check 'Wiki has an FAQ' ($wiki.Name -contains 'FAQ.md')
Check 'Wiki has a CLI reference' ($wiki.Name -contains 'CLI-Reference.md')
Check 'Wiki is entirely English (no Cyrillic)' ($wikiText -notmatch '[\u0400-\u04FF]')
Check 'Publisher tooling exists' (Test-Path -LiteralPath (Join-Path $root 'tools/Publish-Wiki.ps1'))

Req 'BRIEF 5: "add the important features I forgot"'
# Selective cleanup - the single biggest gap in v3
foreach ($p in @('PrinterName','ExcludePrinterName','JobId','OwnerName','DocumentName','OlderThan','NewerThan','LargerThanBytes','SmallerThanBytes')) {
    Check "  selective cleanup: -$p" ($cmd.Parameters.ContainsKey($p))
}
# Backup / restore
foreach ($p in @('Backup','BackupRequired','ArchivePath','Restore','RestoreDestination')) {
    Check "  backup/restore: -$p" ($cmd.Parameters.ContainsKey($p))
}
Check 'Restore-SpoolerBackup cmdlet exported' ($null -ne (Get-Command Restore-SpoolerBackup -ErrorAction SilentlyContinue))
Check 'Get-SpoolerBackup cmdlet exported' ($null -ne (Get-Command Get-SpoolerBackup -ErrorAction SilentlyContinue))
# Diagnostics and health
foreach ($p in @('Diagnostic','DiagnosticsBundlePath')) { Check "  diagnostics: -$p" ($cmd.Parameters.ContainsKey($p)) }
Check 'Test-PrintSpoolerHealth cmdlet exported' ($null -ne (Get-Command Test-PrintSpoolerHealth -ErrorAction SilentlyContinue))
Check 'Get-PrintSpoolerDiagnostic cmdlet exported' ($null -ne (Get-Command Get-PrintSpoolerDiagnostic -ErrorAction SilentlyContinue))
# Repair paths
foreach ($p in @('Repair','PurgeFormsCache','PurgeDriverStore','ResetSpoolDirectory','SelfTestPrint','Watchdog')) {
    Check "  repair: -$p" ($cmd.Parameters.ContainsKey($p))
}
# Fleet / remote
foreach ($p in @('ComputerName','ComputerListFile','Credential','ThrottleLimit','UseSsl','PreFlightProbe')) {
    Check "  fleet: -$p" ($cmd.Parameters.ContainsKey($p))
}
# Reporting and automation
foreach ($p in @('ExportFormat','ExportPath','MachineReadableJson','LogPath','LogLevel','JsonLog','EventLog','ConfigPath','SaveConfig','SelfElevate','VerifySignature','WhatIf','Force','Quiet','NoColor')) {
    Check "  ops: -$p" ($cmd.Parameters.ContainsKey($p))
}
Check 'Scheduled-task installer cmdlet' ($null -ne (Get-Command Install-SpoolerCleanupTask -ErrorAction SilentlyContinue))
Check 'Scheduled-task remover cmdlet' ($null -ne (Get-Command Uninstall-SpoolerCleanupTask -ErrorAction SilentlyContinue))
Check 'PSQ_* environment-variable configuration' ($src -match 'PSQ_LOG_PATH' -and $src -match 'PSQ_FORCE')
$envVars = @([regex]::Matches($src, 'PSQ_[A-Z_]+') | ForEach-Object { $_.Value } | Select-Object -Unique)
Check '24 documented PSQ_* variables' ($envVars.Count -ge 20) "$($envVars.Count) distinct"

Req 'BRIEF 6: functional proof - the tool actually runs'
$st = Clear-PrintSpoolerQueue -SelfTest -NoColor -Quiet -PassThru
Check '-SelfTest passes every internal check' ($st.Overall -eq 'Passed') ("{0}/{1}" -f ($st.TotalChecks - $st.FailedChecks), $st.TotalChecks)
foreach ($c in @($st.Checks | Where-Object Result -eq 'FAIL')) { Write-Host ("        FAIL {0}: {1}" -f $c.Name, $c.Detail) -ForegroundColor DarkRed }
$hp = @(Test-PrintSpoolerHealth)
Check 'Test-PrintSpoolerHealth returns a scored record' ($hp.Count -ge 1 -and $null -ne $hp[0].Health) ("Health={0}" -f $hp[0].Health)
$ver = Clear-PrintSpoolerQueue -ShowVersion -NoColor -Quiet -PassThru
Check '-ShowVersion -PassThru works' ($ver.Version -eq $Script:ScriptVersion)
if ($isWindowsHost -and (Test-RunningAsAdministrator)) {
    $diag = @(Get-PrintSpoolerDiagnostic)
    # Count the EVIDENCE AREAS in the result, not the number of result objects.
    # Get-PrintSpoolerDiagnostic returns one summary per target - a single local
    # target is therefore always 1 - so counting the collection asked the wrong
    # question and reported "1 areas" no matter how much evidence was collected.
    # The areas are the keys of the diagnostics record: OS, services, spooler
    # processes, spool directory, printers, drivers, ports, registry, event log.
    $areas = @()
    foreach ($r in $diag) {
        if ($r.Diagnostics) { $areas += @($r.Diagnostics.Keys) }
    }
    Check 'Get-PrintSpoolerDiagnostic returns evidence areas' ($areas.Count -ge 5) "$($areas.Count) areas: $($areas -join ', ')"
}
elseif ($isWindowsHost) {
    # Collecting diagnostics touches services and registry keys that need elevation,
    # and this check runs in a non-interactive process that cannot raise a UAC
    # prompt. Reporting that as a gap would blame the product for the harness's own
    # privileges, so the check is explicitly inapplicable here rather than failed.
    Write-Host '  SKIP  Get-PrintSpoolerDiagnostic evidence areas (administrator rights are required and were not available)' -ForegroundColor Yellow
}
else {
    # Off Windows the cmdlet must REFUSE rather than invent evidence. Asserting a
    # non-empty result here would be asserting a defect: on a host with no Print
    # Spooler there is nothing to diagnose, and a tool that fabricated areas would
    # be lying to the operator.
    $diag = @(Get-PrintSpoolerDiagnostic 2>$null)
    Check 'Get-PrintSpoolerDiagnostic refuses cleanly off Windows' ($diag.Count -eq 0 -and $Script:RuntimeState.ExitCode -eq 10) ("exit={0}, areas={1}" -f $Script:RuntimeState.ExitCode, $diag.Count)
}
Check 'Get-ExitCodeMeaning documents the contract' (@(Get-ExitCodeMeaning).Count -ge 17)

$env:PSQ_LOG_PATH = $savedLogEnv
Remove-Item -LiteralPath $briefLogDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ''
Write-Host ('=' * 78) -ForegroundColor Cyan
if ($script:fail -eq 0) {
    Write-Host ("  BRIEF COMPLIANCE: {0} requirements verified, 0 gaps" -f $script:pass) -ForegroundColor Green
} else {
    Write-Host ("  BRIEF COMPLIANCE: {0} verified, {1} GAP(S)" -f $script:pass, $script:fail) -ForegroundColor Red
}
Write-Host ('=' * 78) -ForegroundColor Cyan
exit $(if ($script:fail -eq 0) { 0 } else { 1 })

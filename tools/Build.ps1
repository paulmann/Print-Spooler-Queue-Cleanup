#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Builds the Print Spooler Queue Cleanup distributables from src/.

.DESCRIPTION
    The repository keeps the product as small, reviewable, individually testable
    source files under src/. Shipping two hand-maintained copies of the same
    logic (a standalone script and a module) is how the v3 line drifted out of
    sync with its own README. This build assembles BOTH from the same sources,
    so they cannot disagree:

      dist/Clear-PrintSpoolerQueue.ps1                       standalone script
      dist/Clear-PrintSpoolerQueue.bat                       launcher (version stamped)
      dist/PrintSpoolerCleanup/PrintSpoolerCleanup.psd1      module manifest
      dist/PrintSpoolerCleanup/PrintSpoolerCleanup.psm1      module script
      dist/PrintSpoolerCleanup-<version>.zip                 release archive
      dist/checksums.txt                                     SHA-256 for every artifact
      dist/RELEASE-NOTES.md                                  generated release body

    Every generated file is parse-validated before it is accepted, and the build
    fails if any @@TOKEN@@ survived substitution.

.PARAMETER Root
    Repository root. Defaults to the parent of this script's directory.

.PARAMETER Version
    Override the version read from src/Shared/00-Constants.ps1.

.PARAMETER SkipTests
    Do not run tests/Run-Tests.ps1 after building.

.PARAMETER NoZip
    Produce loose files only (faster iteration).

.PARAMETER SigningThumbprint
    Authenticode-sign the generated .ps1/.psm1 with this certificate thumbprint
    (Windows only, requires a code-signing certificate in the user store).

.PARAMETER ReleaseNotesPath
    Path to a CHANGELOG to extract the release body from.

.EXAMPLE
    ./tools/Build.ps1
.EXAMPLE
    ./tools/Build.ps1 -Version 4.0.1 -SigningThumbprint 1A2B3C...
.EXAMPLE
    ./tools/Build.ps1 -SkipTests -NoZip
#>
[CmdletBinding()]
param(
    [string]$Root,
    [string]$Version,
    [string]$OutputPath,
    [switch]$SkipTests,
    [switch]$NoZip,
    [string]$SigningThumbprint,
    [string]$ReleaseNotesPath,
    [switch]$Clean
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

# The module initialiser resolves a log target as soon as the manifest is tested,
# which would otherwise write session logs INTO the artifact directory. Redirect
# logging to a scratch directory for the whole build and restore it afterwards.
$savedLogEnv = $env:PSQ_LOG_PATH
$buildLogDir = Join-Path ([IO.Path]::GetTempPath()) ("psq-build-{0}" -f [Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $buildLogDir -Force
$env:PSQ_LOG_PATH = Join-Path $buildLogDir 'build.log'

if ([string]::IsNullOrWhiteSpace($Root)) { $Root = Split-Path -Parent $PSScriptRoot }
$Root = (Resolve-Path -LiteralPath $Root).Path
$src = Join-Path $Root 'src'
if (-not (Test-Path -LiteralPath $src)) { throw "Source tree not found: $src" }

if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = Join-Path $Root 'dist' }
if ($Clean -and (Test-Path -LiteralPath $OutputPath)) {
    Remove-Item -LiteralPath $OutputPath -Recurse -Force -ErrorAction SilentlyContinue
}
$null = New-Item -ItemType Directory -Path $OutputPath -Force

function Write-Step  { param([string]$T) Write-Host "`n== $T" -ForegroundColor Cyan }
function Write-Ok    { param([string]$T) Write-Host "   OK   $T" -ForegroundColor Green }
function Write-Warn2 { param([string]$T) Write-Host "   WARN $T" -ForegroundColor Yellow }
function Write-Err2  { param([string]$T) Write-Host "   FAIL $T" -ForegroundColor Red }

# ---------------------------------------------------------------------------
# 1. Discover sources and read the authoritative version
# ---------------------------------------------------------------------------
Write-Step 'Discovering sources'
$constantsPath = Join-Path $src 'Shared/00-Constants.ps1'
if (-not (Test-Path -LiteralPath $constantsPath)) { throw "Missing $constantsPath" }
$constantsText = Get-Content -LiteralPath $constantsPath -Raw

if ([string]::IsNullOrWhiteSpace($Version)) {
    if ($constantsText -notmatch "(?m)^\`$Script:ScriptVersion\s*=\s*'([^']+)'") {
        throw 'Could not read $Script:ScriptVersion from the constants file.'
    }
    $Version = $Matches[1]
}
$buildStamp = (Get-Date).ToUniversalTime().ToString('yyyyMMdd.HHmm')
Write-Ok "version $Version (build $buildStamp)"

$orderedFiles = @(
    @{ Path = 'Shared/00-Constants.ps1';    Region = 'CONSTANTS' }
    @{ Path = 'Shared/10-RuntimeState.ps1'; Region = 'RUNTIME STATE' }
    @{ Path = 'Private/01-Platform.ps1';    Region = 'PLATFORM DETECTION' }
    @{ Path = 'Private/02-Logging.ps1';     Region = 'LOGGING' }
    @{ Path = 'Private/03-Console.ps1';     Region = 'CONSOLE RENDERING' }
    @{ Path = 'Private/04-Utility.ps1';     Region = 'UTILITIES' }
    @{ Path = 'Private/05-Configuration.ps1'; Region = 'CONFIGURATION' }
    @{ Path = 'Private/06-RemotePayload.ps1'; Region = 'EXECUTION PAYLOAD TEMPLATE' }
    @{ Path = 'Private/07-Orchestration.ps1'; Region = 'ORCHESTRATION' }
    @{ Path = 'Private/08-BackupRestore.ps1'; Region = 'BACKUP AND RESTORE' }
    @{ Path = 'Private/09-Update.ps1';      Region = 'UPDATE SUBSYSTEM' }
    @{ Path = 'Private/10-Report.ps1';      Region = 'REPORTING AND EXPORT' }
    @{ Path = 'Private/11-Security.ps1';    Region = 'ELEVATION AND INTEGRITY' }
    @{ Path = 'Private/12-ScheduledTask.ps1'; Region = 'TASK SCHEDULER INTEGRATION' }
    @{ Path = 'Private/13-HelpText.ps1';    Region = 'BUILT-IN REFERENCE' }
    @{ Path = 'Private/14-Execution.ps1';   Region = 'WATCHDOG AND SELF-TEST' }
    @{ Path = 'Public/Public-Cmdlets.ps1';  Region = 'PUBLIC CMDLETS' }
    @{ Path = 'Public/Main-Function.ps1';   Region = 'ENTRY POINT' }
)

$parts = New-Object System.Collections.Generic.List[string]
foreach ($f in $orderedFiles) {
    $p = Join-Path $src $f.Path
    if (-not (Test-Path -LiteralPath $p)) { throw "Missing source file: $p" }
    $body = (Get-Content -LiteralPath $p -Raw).TrimEnd()
    $parts.Add("#region $($f.Region)")
    $parts.Add($body)
    $parts.Add("#endregion")
    $parts.Add('')
}
$assembled = ($parts -join "`r`n")

# Build-time substitutions. Everything else that looks like @@TOKEN@@ inside the
# sources belongs to the REMOTE PAYLOAD template and is substituted at RUNTIME by
# Resolve-PayloadScriptBlock (the payload must arrive on the target as literal
# text, so it cannot be expanded at build time).
$assembled = $assembled.Replace('@@BUILD_STAMP@@', $buildStamp)
$assembled = $assembled.Replace('@@SCRIPT_VERSION@@', $Version)
Write-Ok "$($orderedFiles.Count) source files assembled ($([Math]::Round($assembled.Length / 1KB, 1)) KB)"

# ---------------------------------------------------------------------------
# 2. Header and epilogue
# ---------------------------------------------------------------------------
Write-Step 'Rendering header and epilogue'

$functionNames = @(
    'Clear-PrintSpoolerQueue',
    'Get-SpoolerCleanupHelp',
    'Get-SpoolerCleanupVersion',
    'Test-PrintSpoolerHealth',
    'Get-PrintSpoolerDiagnostic',
    'Get-SpoolerCleanupUpdate',
    'Update-SpoolerCleanup',
    'Restore-SpoolerBackup',
    'Get-SpoolerBackup',
    'Install-SpoolerCleanupTask',
    'Uninstall-SpoolerCleanupTask',
    'Get-SpoolerCleanupTask',
    'Start-SpoolerCleanupTask',
    'Get-ExitCodeMeaning'
)

$header = @"
<#
================================================================================
  PRINT SPOOLER QUEUE CLEANUP  -  v$Version
================================================================================

  Clears stuck Windows print queues safely, selectively and verifiably -
  locally or across a fleet - with backup, diagnostics, repair, structured
  logging, machine-readable reports and built-in self-updating.

  WHAT IT DOES
    1. Stops the Print Spooler's dependent services first (Fax, PrintNotify,
       PrintWorkflowUserSvc_*), which is what actually lets the stop succeed.
    2. Inventories the queue from Win32_PrintJob, falling back to .SHD header
       parsing when the Spooler is too hung to answer WMI.
    3. Selects what to remove - by printer, job id, owner, document name, age
       or size - so a busy shared queue keeps its healthy jobs.
    4. Archives the selected files (ZIP + JSON manifest + SHA-256) before
       deleting anything.
    5. Deletes *.SPL/*.SHD with exponential-backoff retries, classifying files
       that are still locked instead of aborting on the first one.
    6. Restarts the Spooler and only the dependents that were running, then
       verifies the queue did not immediately refill.
    7. Reports a structured result and a stable exit code.

  THREE WAYS TO RUN IT
    .\Clear-PrintSpoolerQueue.ps1 -Force -Backup          standalone script
    .\Clear-PrintSpoolerQueue.bat /Force /Backup          double-clickable launcher
    Import-Module .\PrintSpoolerCleanup                   importable module

  START HERE
    .\Clear-PrintSpoolerQueue.ps1 -WhatIf                 see the plan, change nothing
    .\Clear-PrintSpoolerQueue.ps1 -Diagnostic             read-only deep inspection
    .\Clear-PrintSpoolerQueue.ps1 -Help                   the complete reference
    .\Clear-PrintSpoolerQueue.ps1 -CheckUpdate            is there a newer version?

  REQUIREMENTS
    Windows with the Print Spooler feature.
    Windows PowerShell 5.1 or PowerShell 7.x (7.x required for -ThrottleLimit > 1).
    Administrator rights for the local machine; WinRM plus authorization for
    remote targets.

  AUTHOR   Mikhail Deynekin <Mikhail@Deynekin.com>   https://Deynekin.com
  SOURCE   https://github.com/paulmann/Print-Spooler-Queue-Cleanup
  LICENSE  GPL-3.0-or-later
  BUILD    $buildStamp

  This file is generated by tools/Build.ps1 from the sources in src/.
  Do not edit it directly - your changes will be overwritten by the next build.
================================================================================
#>

"@

$epilogue = @"

#region SCRIPT ENTRY POINT
# ----------------------------------------------------------------------------
# Publishes the public functions for dot-sourcing, then runs the cleanup when the
# file was executed (as opposed to dot-sourced or imported).
#
# The exit-code handling here is deliberate. `$MyInvocation is captured IN SCRIPT
# SCOPE, because inside a function it describes the function call rather than the
# script invocation - and getting that wrong is what made v3 report success after
# a failed run.
# ----------------------------------------------------------------------------
`$Script:InvocationName = `$MyInvocation.InvocationName
`$Script:CommandPath    = `$MyInvocation.MyCommand.Path

# Runtime state must exist before ANY function is called, including by a caller
# that only dot-sourced the file. Every helper reads `$Script:RuntimeState, so
# leaving it null until the first Clear-PrintSpoolerQueue call would make
# Get-SpoolerCleanupHelp / Get-SpoolerCleanupVersion throw on a fresh session.
`$null = Initialize-RuntimeState

if (`$Script:InvocationName -eq '.') {
    # Dot-sourced: load the functions into the caller's session and stop there.
    # Running a destructive cleanup as a side effect of dot-sourcing would be
    # unacceptable, so the caller must invoke Clear-PrintSpoolerQueue explicitly.
    Write-Verbose 'Print Spooler Queue Cleanup functions loaded. Call Clear-PrintSpoolerQueue to run a cleanup, or Get-SpoolerCleanupHelp for the reference.'
}
else {
    # Executed directly (.\\script.ps1 or powershell -File script.ps1): run and
    # forward every argument verbatim.
    #
    # Defect 47: the catch block MUST own the final exit code. A parameter
    # binding error means the function body never ran, so RuntimeState.ExitCode
    # is still its initialised "Success" - reading it after the catch (as an
    # earlier revision did) overwrote the 3 back to 0 and every malformed
    # invocation exited "successfully". The catch records its code in a local
    # and the RuntimeState read only happens when there was no binding failure.
    `$scriptExitCode = `$null
    try {
        Clear-PrintSpoolerQueue @args
    }
    catch {
        # A failure before the function's own handler could run (for example a
        # parameter binding error) must still produce a non-zero exit code:
        # 3 = BadArguments in the exit-code contract.
        Write-Error `$_.Exception.Message
        `$scriptExitCode = 3
    }
    if (`$null -eq `$scriptExitCode) {
        `$scriptExitCode = 0
        if (`$Script:RuntimeState -and `$Script:RuntimeState.Contains('ExitCode')) { `$scriptExitCode = [int]`$Script:RuntimeState.ExitCode }
    }
    `$global:LASTEXITCODE = `$scriptExitCode
    exit `$scriptExitCode
}
#endregion
"@

# ---------------------------------------------------------------------------
# 3. Emit the standalone script
# ---------------------------------------------------------------------------
Write-Step 'Emitting dist/Clear-PrintSpoolerQueue.ps1'
$scriptText = $header + $assembled + $epilogue
$scriptOut = Join-Path $OutputPath 'Clear-PrintSpoolerQueue.ps1'
[IO.File]::WriteAllText($scriptOut, $scriptText, (New-Object System.Text.UTF8Encoding($true)))
Write-Ok "$([Math]::Round((Get-Item -LiteralPath $scriptOut).Length / 1KB, 1)) KB"

# ---------------------------------------------------------------------------
# 4. Emit the module
# ---------------------------------------------------------------------------
Write-Step 'Emitting the importable module'
$moduleDir = Join-Path $OutputPath 'PrintSpoolerCleanup'
$null = New-Item -ItemType Directory -Path $moduleDir -Force

$moduleHeader = @"
#
# PrintSpoolerCleanup.psm1 - generated by tools/Build.ps1 from src/. Do not edit.
# Version $Version  |  Build $buildStamp  |  GPL-3.0-or-later
#
# Import-Module .\PrintSpoolerCleanup
# Get-Command -Module PrintSpoolerCleanup
# Get-SpoolerCleanupHelp
#

"@
$moduleText = $moduleHeader + $assembled + @"

# ---------------------------------------------------------------------------
# Module initialisation
# ---------------------------------------------------------------------------
`$null = Initialize-RuntimeState -ScriptDirectory `$PSScriptRoot
Export-ModuleMember -Function @($(($functionNames | ForEach-Object { "'$_'" }) -join ', '))
"@
$moduleOut = Join-Path $moduleDir 'PrintSpoolerCleanup.psm1'
[IO.File]::WriteAllText($moduleOut, $moduleText, (New-Object System.Text.UTF8Encoding($true)))
Write-Ok "PrintSpoolerCleanup.psm1 ($([Math]::Round((Get-Item -LiteralPath $moduleOut).Length / 1KB, 1)) KB)"

# ---- manifest (written by hand rather than via New-ModuleManifest so the
# ---- output is deterministic and identical on every platform)
$guid = 'a7f3c1d2-9b64-4e08-8c2f-51d0e6a4b7c9'
$exportsList = ($functionNames | ForEach-Object { "    '$_'" }) -join ",`r`n"
$releaseNotes = "See CHANGELOG.md for v$Version."
if ($ReleaseNotesPath -and (Test-Path -LiteralPath $ReleaseNotesPath)) {
    $cl = Get-Content -LiteralPath $ReleaseNotesPath -Raw
    if ($cl -match "(?s)##\s*\[?$([regex]::Escape($Version))\]?.*?(?=\r?\n##\s|\z)") {
        $releaseNotes = ($Matches[0].Trim() -replace "'", "''")
        if ($releaseNotes.Length -gt 1500) { $releaseNotes = $releaseNotes.Substring(0, 1500) + ' ...' }
    }
}

$manifest = @"
@{
    # Module manifest for PrintSpoolerCleanup - generated by tools/Build.ps1.

    RootModule             = 'PrintSpoolerCleanup.psm1'
    ModuleVersion          = '$Version'
    CompatiblePSEditions   = @('Desktop', 'Core')
    GUID                   = '$guid'
    Author                 = 'Mikhail Deynekin'
    CompanyName            = 'Deynekin.com'
    Copyright              = '(c) Mikhail Deynekin. All rights reserved.'
    Description            = 'Enterprise-grade Windows Print Spooler queue cleanup: selective job removal, dependent-service orchestration, backup and restore, diagnostics and repair, fleet-wide parallel execution over WinRM, structured logging, machine-readable reports, scheduled-task installation and built-in self-updating from GitHub Releases.'
    PowerShellVersion      = '5.1'
    FunctionsToExport      = @(
$exportsList
    )
    CmdletsToExport        = @()
    VariablesToExport      = @()
    AliasesToExport        = @()

    PrivateData = @{
        PSData = @{
            Tags         = @('Windows', 'PrintSpooler', 'Spooler', 'PrintQueue', 'Printing', 'Maintenance', 'HelpDesk', 'SysAdmin', 'WinRM', 'Remote', 'Server', 'Cleanup', 'Diagnostics', 'Automation', 'TaskScheduler')
            LicenseUri   = 'https://github.com/paulmann/Print-Spooler-Queue-Cleanup/blob/main/LICENSE'
            ProjectUri   = 'https://github.com/paulmann/Print-Spooler-Queue-Cleanup'
            ReleaseNotes = @'
$releaseNotes
'@
        }
    }
}
"@
$manifestOut = Join-Path $moduleDir 'PrintSpoolerCleanup.psd1'
[IO.File]::WriteAllText($manifestOut, $manifest, (New-Object System.Text.UTF8Encoding($true)))
Write-Ok 'PrintSpoolerCleanup.psd1'

# ---------------------------------------------------------------------------
# 5. Stamp the batch launcher
# ---------------------------------------------------------------------------
Write-Step 'Emitting dist/Clear-PrintSpoolerQueue.bat'
$batTemplate = Join-Path $Root 'bat/Clear-PrintSpoolerQueue.bat'
if (-not (Test-Path -LiteralPath $batTemplate)) { throw "Missing $batTemplate" }
$batText = Get-Content -LiteralPath $batTemplate -Raw
$batText = $batText.Replace('@@SCRIPT_VERSION@@', $Version)
$batOut = Join-Path $OutputPath 'Clear-PrintSpoolerQueue.bat'
# CRLF + ASCII: a .bat with LF-only endings misbehaves in cmd.exe, and non-ASCII
# bytes render as mojibake in the OEM code page.
$batText = $batText -replace "`r`n", "`n" -replace "`n", "`r`n"
[IO.File]::WriteAllText($batOut, $batText, (New-Object System.Text.ASCIIEncoding))
Write-Ok "Clear-PrintSpoolerQueue.bat ($([Math]::Round((Get-Item -LiteralPath $batOut).Length / 1KB, 1)) KB)"

# ---------------------------------------------------------------------------
# 6. Copy the supporting files a release needs
# ---------------------------------------------------------------------------
Write-Step 'Copying supporting files'
$copies = @(
    @{ From = 'config/PrintSpoolerCleanup.config.psd1'; To = 'PrintSpoolerCleanup.config.psd1' }
    @{ From = 'README.md';   To = 'README.md' }
    @{ From = 'LICENSE';     To = 'LICENSE' }
    @{ From = 'CHANGELOG.md'; To = 'CHANGELOG.md' }
)
foreach ($c in $copies) {
    $from = Join-Path $Root $c.From
    if (Test-Path -LiteralPath $from) {
        Copy-Item -LiteralPath $from -Destination (Join-Path $OutputPath $c.To) -Force
        Write-Ok $c.To
    }
    else { Write-Warn2 "not found, skipped: $($c.From)" }
}

# ---------------------------------------------------------------------------
# 7. Validate everything that was generated
# ---------------------------------------------------------------------------
Write-Step 'Validating generated files'
$validationFailures = 0

$toParse = @(
    $scriptOut, $moduleOut, $manifestOut,
    (Join-Path $OutputPath 'PrintSpoolerCleanup.config.psd1')
)
foreach ($f in $toParse) {
    if (-not (Test-Path -LiteralPath $f)) { continue }
    $tokens = $null; $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($f, [ref]$tokens, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) {
        $validationFailures++
        Write-Err2 "$(Split-Path -Leaf $f): $($errors.Count) parse error(s)"
        $errors | Select-Object -First 5 | ForEach-Object { Write-Host ("        line {0}: {1}" -f $_.Extent.StartLineNumber, $_.Message) -ForegroundColor DarkRed }
    }
    else { Write-Ok "$(Split-Path -Leaf $f) parses cleanly" }
}

# Token substitution audit.
#
# The .bat and the manifest must contain NO @@TOKEN@@ at all: everything in them
# is build-time. The .ps1/.psm1 legitimately retain the remote-payload tokens,
# because that template is substituted at runtime on the target machine. So for
# those two we assert the exact whitelist instead of "nothing left", which is a
# stronger check: an unknown token means a typo that would break the payload.
$payloadTokenWhitelist = @(
    '@@SPOOLER_SERVICE@@', '@@DEPENDENT_SERVICES@@', '@@SPOOL_EXTENSIONS@@',
    '@@LEGACY_EXTENSIONS@@', '@@DEFAULT_SPOOL_SUBPATH@@', '@@FORMS_SUBPATH@@',
    '@@GUARD_REQUIRED_LEAF@@', '@@GUARD_MIN_DEPTH@@', '@@GUARD_MAX_FILES@@',
    '@@GUARD_MAX_BYTES@@', '@@FORBIDDEN_ROOTS@@', '@@RETRY_BACKOFF@@',
    '@@SCRIPT_NAME@@', '@@SCRIPT_VERSION@@', '@@PRODUCT_FAMILY@@', '@@TOKEN@@'
)

foreach ($f in @($batOut, $manifestOut)) {
    if (-not (Test-Path -LiteralPath $f)) { continue }
    $text = Get-Content -LiteralPath $f -Raw
    $leftover = @([regex]::Matches($text, '@@[A-Z_]+@@') | ForEach-Object { $_.Value } | Select-Object -Unique)
    if ($leftover.Count -gt 0) {
        $validationFailures++
        Write-Err2 "$(Split-Path -Leaf $f) still contains build tokens: $($leftover -join ', ')"
    }
    else { Write-Ok "$(Split-Path -Leaf $f) has no leftover tokens" }
}

foreach ($f in @($scriptOut, $moduleOut)) {
    if (-not (Test-Path -LiteralPath $f)) { continue }
    $text = Get-Content -LiteralPath $f -Raw
    $found = @([regex]::Matches($text, '@@[A-Z_]+@@') | ForEach-Object { $_.Value } | Select-Object -Unique)
    $unknown = @($found | Where-Object { $payloadTokenWhitelist -notcontains $_ })
    $buildTime = @($found | Where-Object { $_ -eq '@@BUILD_STAMP@@' })
    if ($unknown.Count -gt 0 -or $buildTime.Count -gt 0) {
        $validationFailures++
        Write-Err2 "$(Split-Path -Leaf $f) has unexpected tokens: $(@($unknown + $buildTime) -join ', ')"
    }
    else {
        $expected = @($payloadTokenWhitelist | Where-Object { $found -contains $_ -and $_ -ne '@@TOKEN@@' })
        Write-Ok "$(Split-Path -Leaf $f) retains $($expected.Count) payload token(s) for runtime substitution"
        # The payload template must be present, or the tokens mean nothing.
        if ($text -notmatch [regex]::Escape('$Script:RemotePayloadTemplate')) {
            $validationFailures++
            Write-Err2 "$(Split-Path -Leaf $f) is missing the remote payload template"
        }
    }
}

# Line-ending sanity for the .bat
$batBytes = [IO.File]::ReadAllBytes($batOut)
$crlfCount = 0
for ($i = 1; $i -lt $batBytes.Length; $i++) { if ($batBytes[$i] -eq 10 -and $batBytes[$i - 1] -eq 13) { $crlfCount++ } }
if ($crlfCount -lt 50) { $validationFailures++; Write-Err2 "Clear-PrintSpoolerQueue.bat does not have CRLF endings ($crlfCount found)" }
else { Write-Ok "batch line endings are CRLF ($crlfCount)" }
$nonAscii = @($batBytes | Where-Object { $_ -gt 127 }).Count
if ($nonAscii -gt 0) { $validationFailures++; Write-Err2 "batch file contains $nonAscii non-ASCII byte(s)" }
else { Write-Ok 'batch file is pure ASCII' }

# The manifest must be loadable and must export what we promised.
try {
    $testModule = Test-ModuleManifest -Path $manifestOut -ErrorAction Stop
    $exported = @($testModule.ExportedFunctions.Keys)
    $missing = @($functionNames | Where-Object { $exported -notcontains $_ })
    if ($missing.Count -gt 0) {
        Write-Warn2 "manifest does not export: $($missing -join ', ')"
    }
    else { Write-Ok "manifest exports all $($functionNames.Count) functions" }
}
catch {
    # Test-ModuleManifest can fail on non-Windows for reasons unrelated to the
    # manifest content; report but do not fail the build on it.
    Write-Warn2 "Test-ModuleManifest could not run here: $($_.Exception.Message)"
}

if ($validationFailures -gt 0) { throw "Build validation failed with $validationFailures problem(s)." }

# ---------------------------------------------------------------------------
# 8. Optional code signing
# ---------------------------------------------------------------------------
if ($SigningThumbprint) {
    Write-Step 'Authenticode signing'
    if (-not (Get-Command -Name Set-AuthenticodeSignature -ErrorAction SilentlyContinue)) {
        Write-Warn2 'Set-AuthenticodeSignature is unavailable on this platform; signing skipped.'
    }
    else {
        foreach ($f in @($scriptOut, $moduleOut)) {
            try {
                $sig = Set-AuthenticodeSignature -FilePath $f -Certificate (Get-Item "Cert:\CurrentUser\My\$SigningThumbprint") -TimestampServer 'http://timestamp.digicert.com' -ErrorAction Stop
                if ($sig.Status -eq 'Valid') { Write-Ok "signed $(Split-Path -Leaf $f)" }
                else { Write-Warn2 "signing $(Split-Path -Leaf $f) returned $($sig.Status)" }
            }
            catch { Write-Warn2 "could not sign $(Split-Path -Leaf $f): $($_.Exception.Message)" }
        }
    }
}

# ---------------------------------------------------------------------------
# 9. Release notes
# ---------------------------------------------------------------------------
Write-Step 'Generating release notes'
$notesPath = Join-Path $OutputPath 'RELEASE-NOTES.md'
$notes = New-Object System.Text.StringBuilder
[void]$notes.AppendLine("# Print Spooler Queue Cleanup v$Version")
[void]$notes.AppendLine('')
[void]$notes.AppendLine("Build $buildStamp | PowerShell 5.1 and 7.x | Windows 10/11/Server | GPL-3.0-or-later")
[void]$notes.AppendLine('')
[void]$notes.AppendLine('## Artifacts')
[void]$notes.AppendLine('')
[void]$notes.AppendLine('| File | Purpose |')
[void]$notes.AppendLine('|---|---|')
[void]$notes.AppendLine('| `Clear-PrintSpoolerQueue.ps1` | Standalone script. Copy this one file and run it. |')
[void]$notes.AppendLine('| `Clear-PrintSpoolerQueue.bat` | Double-clickable launcher. Detects PowerShell 7, can self-elevate. |')
[void]$notes.AppendLine('| `PrintSpoolerCleanup-*.zip` | Importable PowerShell module + script + launcher + config template. |')
[void]$notes.AppendLine('| `PrintSpoolerCleanup.config.psd1` | Documented configuration template. |')
[void]$notes.AppendLine('| `checksums.txt` | SHA-256 for every artifact. The updater verifies against this. |')
[void]$notes.AppendLine('')
if ($ReleaseNotesPath -and (Test-Path -LiteralPath $ReleaseNotesPath)) {
    $cl = Get-Content -LiteralPath $ReleaseNotesPath -Raw
    if ($cl -match "(?s)(##\s*\[?$([regex]::Escape($Version))\]?.*?)(?=\r?\n##\s|\z)") {
        [void]$notes.AppendLine($Matches[1].Trim())
        [void]$notes.AppendLine('')
    }
}
[void]$notes.AppendLine('## Quick start')
[void]$notes.AppendLine('')
[void]$notes.AppendLine('```powershell')
[void]$notes.AppendLine('.\Clear-PrintSpoolerQueue.ps1 -WhatIf        # see the plan, change nothing')
[void]$notes.AppendLine('.\Clear-PrintSpoolerQueue.ps1 -Diagnostic    # read-only deep inspection')
[void]$notes.AppendLine('.\Clear-PrintSpoolerQueue.ps1 -Force -Backup # fix it, with an archive first')
[void]$notes.AppendLine('.\Clear-PrintSpoolerQueue.ps1 -Help          # the complete reference')
[void]$notes.AppendLine('```')
[void]$notes.AppendLine('')
[void]$notes.AppendLine('## Verification')
[void]$notes.AppendLine('')
[void]$notes.AppendLine('```powershell')
[void]$notes.AppendLine('Get-FileHash .\Clear-PrintSpoolerQueue.ps1 -Algorithm SHA256')
[void]$notes.AppendLine('Get-Content .\checksums.txt')
[void]$notes.AppendLine('```')
[void]$notes.AppendLine('')
[void]$notes.AppendLine("Full documentation: https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki")
Set-Content -LiteralPath $notesPath -Value $notes.ToString() -Encoding UTF8
Write-Ok 'RELEASE-NOTES.md'

# ---------------------------------------------------------------------------
# 10. ZIP + checksums
# ---------------------------------------------------------------------------
if (-not $NoZip) {
    Write-Step 'Packaging'
    $zipName = "{0}-{1}.zip" -f 'PrintSpoolerCleanup', $Version
    $zipPath = Join-Path $OutputPath $zipName
    if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }

    $staging = Join-Path ([IO.Path]::GetTempPath()) ("psq-package-{0}" -f [Guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $staging -Force
    try {
        Copy-Item -LiteralPath $scriptOut -Destination $staging -Force
        Copy-Item -LiteralPath $batOut -Destination $staging -Force
        Copy-Item -LiteralPath $moduleDir -Destination $staging -Recurse -Force
        Get-ChildItem -LiteralPath (Join-Path $staging (Split-Path -Leaf $moduleDir)) -Recurse -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -eq 'Logs' } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        foreach ($extra in @('PrintSpoolerCleanup.config.psd1', 'README.md', 'LICENSE', 'CHANGELOG.md')) {
            $p = Join-Path $OutputPath $extra
            if (Test-Path -LiteralPath $p) { Copy-Item -LiteralPath $p -Destination $staging -Force }
        }
        if (Get-Command -Name Compress-Archive -ErrorAction SilentlyContinue) {
            Compress-Archive -Path (Join-Path $staging '*') -DestinationPath $zipPath -CompressionLevel Optimal -ErrorAction Stop
        }
        else {
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [System.IO.Compression.ZipFile]::CreateFromDirectory($staging, $zipPath)
        }
        Write-Ok "$zipName ($([Math]::Round((Get-Item -LiteralPath $zipPath).Length / 1KB, 1)) KB)"
    }
    finally {
        Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Step 'Computing checksums'
$checksumFiles = @(Get-ChildItem -LiteralPath $OutputPath -File | Where-Object { $_.Name -ne 'checksums.txt' -and $_.Extension -notin @('.md') } | Sort-Object Name)
$checksumText = ($checksumFiles | ForEach-Object {
    $h = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
    "{0}  {1}" -f $h, $_.Name
}) -join "`n"
$checksumPath = Join-Path $OutputPath 'checksums.txt'
[IO.File]::WriteAllText($checksumPath, $checksumText + "`n", (New-Object System.Text.UTF8Encoding($false)))
Write-Ok "checksums.txt ($($checksumFiles.Count) files)"

# ---------------------------------------------------------------------------
# 11. Tests
# ---------------------------------------------------------------------------
if (-not $SkipTests) {
    Write-Step 'Running the test suite'
    $testRunner = Join-Path $Root 'tests/Run-Tests.ps1'
    if (Test-Path -LiteralPath $testRunner) {
        & $testRunner -Root $Root
        if ($LASTEXITCODE -ne 0) { throw "Test suite failed (exit code $LASTEXITCODE)." }
        Write-Ok 'all tests passed'
    }
    else { Write-Warn2 "no test runner at $testRunner" }
}

Write-Step 'Build complete'
Get-ChildItem -LiteralPath $OutputPath -Recurse -File |
    Where-Object { $_.FullName -notmatch '[\\/]Logs[\\/]' } |
    Sort-Object FullName | ForEach-Object {
    Write-Host ("   {0,-46} {1,10} KB" -f $_.FullName.Substring($OutputPath.Length).TrimStart([IO.Path]::DirectorySeparatorChar, '/'), ($_.Length / 1KB).ToString('N1', [Globalization.CultureInfo]::InvariantCulture))
}
Write-Host ''
Write-Host "Artifacts are in $OutputPath" -ForegroundColor Green

# An explicit exit code matters: without it, `$LASTEXITCODE in a caller reflects
# whatever external command ran last inside this script, and a green build can be
# reported as a failure (this exact bug made the test gate abort at stage 1).
$env:PSQ_LOG_PATH = $savedLogEnv
Remove-Item -LiteralPath $buildLogDir -Recurse -Force -ErrorAction SilentlyContinue
# A release folder must contain artifacts and nothing else. Importing the module
# or dot-sourcing the script initialises runtime state, which resolves a log
# target; when %ProgramData% is unavailable (Linux CI, or a locked-down host) the
# fallback is the script directory, which would leave a Logs\ folder inside the
# shipped artifacts. Sweep every log directory the build can have produced.
foreach ($junk in @('Logs', '.update')) {
    foreach ($dir in @(Get-ChildItem -LiteralPath $OutputPath -Recurse -Directory -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $junk })) {
        Remove-Item -LiteralPath $dir.FullName -Recurse -Force -ErrorAction SilentlyContinue
    }
}
foreach ($junkFile in @('update-state.json', 'pending-update.json')) {
    foreach ($f in @(Get-ChildItem -LiteralPath $OutputPath -Recurse -File -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $junkFile })) {
        Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue
    }
}
exit 0

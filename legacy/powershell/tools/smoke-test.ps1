<#
.SYNOPSIS
    End-to-end smoke test of the BUILT artifact (dist/Clear-PrintSpoolerQueue.ps1).

.DESCRIPTION
    Complements tests/Run-Tests.ps1, which unit-tests the sources. This script
    exercises the assembled distributable the way an operator would: dot-source
    it, enumerate the public surface, render every help section, run -SelfTest,
    and verify the guards, parsers, config safety, logging, rotation, backup
    round-trip, exports, update discovery and platform handling.

    It runs on Windows and on Linux/macOS. Platform-dependent checks assert the
    correct REFUSAL behaviour off Windows rather than being skipped, so the suite
    is meaningful in CI on any runner.

.EXAMPLE
    ./tools/smoke-test.ps1
#>
[CmdletBinding()]
param(
    [string]$Root,
    [switch]$SkipNetwork
)

$ErrorActionPreference = 'Continue'
if ([string]::IsNullOrWhiteSpace($Root)) { $Root = Split-Path -Parent $PSScriptRoot }
$script = Join-Path $Root 'dist/Clear-PrintSpoolerQueue.ps1'
if (-not (Test-Path -LiteralPath $script)) { throw "Built artifact not found: $script (run tools/Build.ps1 first)" }

$script:failures = 0
$script:passes = 0
function Check {
    param([string]$Name, [bool]$Ok, [string]$Detail = '')
    if ($Ok) { $script:passes++; Write-Host ("  PASS  {0,-46} {1}" -f $Name, $Detail) -ForegroundColor Green }
    else { $script:failures++; Write-Host ("  FAIL  {0,-46} {1}" -f $Name, $Detail) -ForegroundColor Red }
}
function Section { param([string]$Title) Write-Host ''; Write-Host "=== $Title ===" -ForegroundColor Cyan }

. $script

Section '1. Dot-source the built artifact'
Check 'Dot-sourcing does not execute a cleanup' $true
Check 'RuntimeState was initialised' ($null -ne $Script:RuntimeState)

Section '2. Public surface'
$expected = @('Clear-PrintSpoolerQueue','Get-SpoolerCleanupHelp','Get-SpoolerCleanupVersion','Test-PrintSpoolerHealth','Get-PrintSpoolerDiagnostic','Get-SpoolerCleanupUpdate','Update-SpoolerCleanup','Restore-SpoolerBackup','Get-SpoolerBackup','Install-SpoolerCleanupTask','Uninstall-SpoolerCleanupTask','Get-SpoolerCleanupTask','Start-SpoolerCleanupTask','Get-ExitCodeMeaning')
$missing = @($expected | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) })
Check "All $($expected.Count) public functions exist" ($missing.Count -eq 0) $(if ($missing) { "missing: $($missing -join ', ')" } else { '' })

Section '3. Parameter contract'
$cmd = Get-Command Clear-PrintSpoolerQueue
$required = @('ComputerName','ComputerListFile','Credential','ThrottleLimit','PreFlightProbe','UseSsl','PrinterName','ExcludePrinterName','JobId','DocumentName','OwnerName','OlderThan','NewerThan','LargerThanBytes','SmallerThanBytes','SpoolDirectory','Diagnostic','Repair','Restore','RestoreDestination','CheckUpdate','Update','AutoUpdate','SelfTest','Watchdog','SelfTestPrint','Force','Backup','BackupPath','BackupRequired','ArchivePath','RetryCount','RetryDelayMs','VerifySignature','ExpectedSignerThumbprint','NonInteractive','WatchdogMinutes','PurgeFormsCache','PurgeDriverStore','ResetSpoolDirectory','StopDependentServices','RestartDependentServices','Quiet','NoColor','Ascii','NoProgress','MachineReadableJson','ExportFormat','ExportPath','PassThru','DiagnosticsBundlePath','LogPath','LogLevel','JsonLog','EventLog','LogRetentionDays','NoLogRotation','ConfigPath','SaveConfig','NoConfigFile','ServiceTimeoutSeconds','SettleSeconds','EventLogHours','UpdateChannel','IncludePrerelease','SkipChecksumVerification','GitHubToken','NoUpdateCheck','ApplyUpdateNow','SelfElevate','NoElevate','NoProfile','Help','ListHelpSections','ShowVersion')
$missP = @($required | Where-Object { -not $cmd.Parameters.ContainsKey($_) })
Check "All $($required.Count) parameters exist" ($missP.Count -eq 0) $(if ($missP) { "missing: $($missP -join ', ')" } else { "$($cmd.Parameters.Count) total" })
Check 'SupportsShouldProcess is enabled' (@($cmd.Parameters.Keys) -contains 'WhatIf')
Check 'ComputerName accepts pipeline input' (@($cmd.Parameters['ComputerName'].Attributes | Where-Object { $_.ValueFromPipeline }).Count -gt 0)

Section '4. Help renders (every section)'
foreach ($sec in (Get-HelpSectionNames)) {
    if ($sec -eq 'All') { continue }
    try { $t = Get-HelpText -Section $sec -IncludeParameterMetadata; Check "Section '$sec'" ($t.Length -gt 200) "$($t.Length) chars" }
    catch { Check "Section '$sec'" $false $_.Exception.Message }
}
try { $all = Get-HelpText -Section 'All' -IncludeParameterMetadata; Check 'Full document' ($all.Length -gt 40000) "$($all.Length) chars" }
catch { Check 'Full document' $false $_.Exception.Message }
try { $f = Get-HelpText -Section 'All' -Search 'backup'; Check 'Search filter works' ($f.Length -gt 100 -and $f.Length -lt $all.Length) "$($f.Length) chars" }
catch { Check 'Search filter works' $false $_.Exception.Message }

Section '5. Help documents every parameter'
$docs = @(Get-ParameterDocumentation)
$skip = @('Verbose','Debug','ErrorAction','WarningAction','InformationAction','ErrorVariable','WarningVariable','InformationVariable','OutVariable','OutBuffer','PipelineVariable','ProgressAction','Confirm','WhatIf')
$undoc = @()
foreach ($p in $cmd.Parameters.Keys) {
    if ($skip -contains $p) { continue }
    if (@($docs | Where-Object { $_.Name -eq $p }).Count -eq 0) { $undoc += $p }
}
Check 'Every parameter is documented' ($undoc.Count -eq 0) $(if ($undoc) { $undoc -join ', ' } else { "$($docs.Count) documented" })
$nodesc = @($docs | Where-Object { -not $_.Description -or $_.Description -match 'No description available' })
Check 'No placeholder descriptions' ($nodesc.Count -eq 0) $(if ($nodesc) { ($nodesc | ForEach-Object { $_.Name }) -join ', ' } else { '' })

Section '6. Built-in self test'
try { $st = Clear-PrintSpoolerQueue -SelfTest -NoColor -PassThru; foreach ($c in $st.Checks) { Check $c.Name ($c.Result -eq 'PASS') $c.Detail } }
catch { Check '-SelfTest runs' $false $_.Exception.Message }

Section '7. Payload generation'
try {
    $sb = Resolve-PayloadScriptBlock -Force   # assigned: must not reach the pipeline
    Check 'Payload scriptblock generated' ($null -ne $sb)
    $txt = $Script:RuntimeState.PayloadText
    $left = @([regex]::Matches($txt, '@@[A-Z_]+@@') | ForEach-Object { $_.Value } | Select-Object -Unique)
    Check 'No leftover payload tokens' ($left.Count -eq 0) $(if ($left) { $left -join ', ' } else { '' })
    Check 'Dependent services substituted' ($txt -match "\`$depNames\s+=\s+@\('PrintWorkflowUserSvc', 'Fax', 'PrintNotify'\)")
    Check 'Extensions substituted' ($txt -match "\`$exts\s+=\s+@\('\.SPL', '\.SHD'\)")
    Check 'Guard leaf substituted' ($txt -match "notlike 'PRINTERS'")
    Check 'Forbidden roots substituted' ($txt -match "'C:\\Windows\\System32'")
    Check 'Backoff multiplier substituted' ($txt -match 'Math\]::Pow\(2,')
    Check 'Payload contains no unquoted token remnants' ($txt -notmatch "@\('@@")
    $pt = $null; $pe = $null
    [void][System.Management.Automation.Language.Parser]::ParseInput($txt, [ref]$pt, [ref]$pe)
    Check 'Generated payload parses' (($pe -eq $null) -or ($pe.Count -eq 0))
} catch { Check 'Payload generation' $false $_.Exception.Message }

Section '8. Exit-code contract'
$codes = @($Script:ExitCode.Values)
Check 'No duplicate exit codes' (@($codes | Group-Object | Where-Object { $_.Count -gt 1 }).Count -eq 0)
Check 'Success is 0' ($Script:ExitCode.Success -eq 0)
Check 'All codes have a meaning' (@(Get-ExitCodeMeaning | Where-Object { -not $_.Description }).Count -eq 0) "$($codes.Count) codes"

Section '9. Version comparison'
$cases = @(@('3.1.0','4.0.0',-1),@('4.0.0','4.0.0',0),@('4.0.1','4.0.0',1),@('4.10.0','4.9.0',1),@('4.0.0-beta','4.0.0',-1),@('v4.0.0','4.0.0',0),@('4.0','4.0.0.0',0),@('','4.0.0',-1),@('4.0.0','',1),@('garbage','4.0.0',-1),@('4.0.0','garbage',1))
foreach ($c in $cases) {
    $got = Compare-ProductVersion -Reference $c[0] -Difference $c[1]
    Check "'$($c[0])' vs '$($c[1])'" ($got -eq $c[2]) "got $got, expected $($c[2])"
}
Check 'Prereleases ignored by default' (-not (Test-NewerVersion -Current '4.0.0' -Candidate '4.1.0-rc1'))
Check 'Prereleases accepted when asked' (Test-NewerVersion -Current '4.0.0' -Candidate '4.1.0-rc1' -IncludePrerelease)
Check 'Older version is not an update' (-not (Test-NewerVersion -Current '4.0.0' -Candidate '3.9.9'))

Section '10. Spool-directory safety guards'
foreach ($bad in @('C:\','C:\Windows','C:\Users','C:\Program Files','\\server\share\PRINTERS','C:\a\b','','C:\Data','C:\PRINTERS','D:\PRINTERS','rel\ative\spool\PRINTERS','C:\Windows\System32','C:\Users\Public\Documents','C:\Users\PRINTERS','D:\PrintSpool\PRINTERS')) {
    $r = Test-SpoolDirectorySafeInternal -Path $bad
    Check "refuses '$bad'" (-not $r.IsSafe) $r.Reason
}
# A relocated spool root is a legitimate administrator choice and MUST be
# accepted; only paths that cannot be a spool directory are refused.
foreach ($good in @('C:\Windows\System32\spool\PRINTERS','D:\Spool\PRINTERS','E:\Print$\spool\PRINTERS','D:\Print\spool\PRINTERS','F:\Very\Deep\spool\PRINTERS')) {
    $r = Test-SpoolDirectorySafeInternal -Path $good
    Check "accepts '$good'" $r.IsSafe $r.Reason
}

Section '11. Value parsers'
Check "'90s' -> 90 seconds" ((ConvertFrom-AnyTimeSpan -Value '90s').TotalSeconds -eq 90)
Check "'2h' -> 2 hours" ((ConvertFrom-AnyTimeSpan -Value '2h').TotalHours -eq 2)
Check "'7d' -> 7 days" ((ConvertFrom-AnyTimeSpan -Value '7d').TotalDays -eq 7)
Check "'1w' -> 7 days" ((ConvertFrom-AnyTimeSpan -Value '1w').TotalDays -eq 7)
Check "'45m' -> 45 minutes" ((ConvertFrom-AnyTimeSpan -Value '45m').TotalMinutes -eq 45)
Check "'01:30:00' -> 90 minutes" ((ConvertFrom-AnyTimeSpan -Value '01:30:00').TotalMinutes -eq 90)
Check '[TimeSpan] passthrough' ((ConvertFrom-AnyTimeSpan -Value ([TimeSpan]::FromMinutes(5))).TotalMinutes -eq 5)
Check 'numeric seconds' ((ConvertFrom-AnyTimeSpan -Value 45).TotalSeconds -eq 45)
$thrown = $false; try { ConvertFrom-AnyTimeSpan -Value 'nonsense' | Out-Null } catch { $thrown = $true }
Check 'garbage throws' $thrown
Check "'10MB' -> 10485760" ((Format-FileSizeArgument -Value '10MB') -eq 10485760)
Check "'1.5GB' -> 1610612736" ((Format-FileSizeArgument -Value '1.5GB') -eq 1610612736)
Check "'512' -> 512" ((Format-FileSizeArgument -Value '512') -eq 512)
Check 'garbage size -> null' ($null -eq (Format-FileSizeArgument -Value 'banana'))
Check 'Format-ByteSize 1536' ((Format-ByteSize 1536) -eq '1.50 KB')
Check 'Format-ByteSize 0' ((Format-ByteSize 0) -eq '0 B')
Check 'ConvertTo-SafeFileName neutralises illegal chars' ((ConvertTo-SafeFileName -Value 'a/b:c*d?') -notmatch '[/:*?]')
Check 'ConvertTo-SafeFileName never returns empty' ((ConvertTo-SafeFileName -Value '///').Length -gt 0)
Check 'ConvertTo-SafeFileName honours MaxLength' ((ConvertTo-SafeFileName -Value ('x' * 300) -MaxLength 20).Length -le 20)

Section '12. Configuration'
$def = Get-DefaultConfiguration
Check 'Defaults validate against their own schema' (@(Test-Configuration -Configuration $def).Count -eq 0)
$bad = Get-DefaultConfiguration; $bad['Service']['TimeoutSeconds'] = 99999
Check 'Out-of-range value rejected' (@(Test-Configuration -Configuration $bad).Count -gt 0)
$bad2 = Get-DefaultConfiguration; $bad2['Nope'] = @{}
Check 'Unknown section rejected' (@(Test-Configuration -Configuration $bad2).Count -gt 0)
$bad3 = Get-DefaultConfiguration; $bad3['Cleanup']['IncludeExtensions'] = @('*')
Check 'Wildcard extension rejected' (@(Test-Configuration -Configuration $bad3).Count -gt 0)
$bad4 = Get-DefaultConfiguration; $bad4['Service']['NotAKey'] = 1
Check 'Unknown key rejected' (@(Test-Configuration -Configuration $bad4).Count -gt 0)
$merged = Merge-Configuration -Base (Get-DefaultConfiguration) -Override @{ Service = @{ TimeoutSeconds = 120 } }
Check 'Deep merge overrides one key' ($merged['Service']['TimeoutSeconds'] -eq 120)
Check 'Deep merge keeps siblings' ($merged['Service']['SettleSeconds'] -eq 2)
Check 'Deep merge does not mutate base' ((Get-DefaultConfiguration)['Service']['TimeoutSeconds'] -eq 60)
$nullMerge = Merge-Configuration -Base (Get-DefaultConfiguration) -Override @{ Service = @{ TimeoutSeconds = $null } }
Check 'Null override does not erase a default' ($nullMerge['Service']['TimeoutSeconds'] -eq 60)

Section '13. Config file safety (AST, never executed)'
$tmp = Join-Path ([IO.Path]::GetTempPath()) ("psq-cfg-{0}.psd1" -f [Guid]::NewGuid().ToString('N'))
try {
    Set-Content -LiteralPath $tmp -Value "@{ Service = @{ TimeoutSeconds = 90 } }" -Encoding UTF8
    $r = ConvertFrom-ConfigFileAst -Path $tmp
    Check 'Benign config parses' ($r.Errors.Count -eq 0 -and $r.Data['Service']['TimeoutSeconds'] -eq 90)
    Set-Content -LiteralPath $tmp -Value "@{ Service = @{ TimeoutSeconds = 90 } }`nRemove-Item C:\ -Recurse -Force" -Encoding UTF8
    $r2 = ConvertFrom-ConfigFileAst -Path $tmp
    Check 'Executable statement rejected' ($r2.Errors.Count -gt 0) $(if ($r2.Errors) { $r2.Errors[0] } else { '' })
    Set-Content -LiteralPath $tmp -Value "@{ Service = @{ TimeoutSeconds = (Invoke-Expression '1+1') } }" -Encoding UTF8
    $r3 = ConvertFrom-ConfigFileAst -Path $tmp
    Check 'Command invocation in a value rejected' ($r3.Errors.Count -gt 0)
    Set-Content -LiteralPath $tmp -Value "@{ broken" -Encoding UTF8
    $r4 = ConvertFrom-ConfigFileAst -Path $tmp
    Check 'Malformed file reports a parse error' ($r4.Errors.Count -gt 0)
}
finally { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }

Section '14. Checksum handling'
$sample = "# comment`nabc123def456abc123def456abc123def456abc123def456abc123def456abc1  Clear-PrintSpoolerQueue.ps1`n0000000000000000000000000000000000000000000000000000000000000000 *PrintSpoolerCleanup-4.0.0.zip`n"
$parsed = ConvertFrom-ChecksumText -Text $sample
Check 'GNU coreutils format parsed' ($parsed.Count -eq 2)
Check 'Binary marker handled' ($parsed.Contains('PrintSpoolerCleanup-4.0.0.zip'))
$p2 = ConvertFrom-ChecksumText -Text "file.ps1: abc123def456abc123def456abc123def456abc123def456abc123def456abc1"
Check 'Reversed format parsed' ($p2.Contains('file.ps1'))
Check 'Comments ignored' (-not ($parsed.Contains('# comment')))
Check 'Hash match is case-insensitive' (Test-HashMatch -Expected 'ABC123' -Actual 'abc123')
Check 'Hash mismatch rejected' (-not (Test-HashMatch -Expected 'ABC123' -Actual 'abc124'))
Check 'Empty hash rejected' (-not (Test-HashMatch -Expected '' -Actual 'abc'))
$round = ConvertTo-ChecksumText -FilePaths @($script)
Check 'Checksum generation round-trips' ((ConvertFrom-ChecksumText -Text $round).Contains('Clear-PrintSpoolerQueue.ps1'))

Section '15. Parameter conflict detection'
Check '-Restore + -Repair' (@(Test-ParameterCombination -Bound @{ Restore = 'x'; Repair = $true }).Count -gt 0)
Check '-Diagnostic + -Repair' (@(Test-ParameterCombination -Bound @{ Diagnostic = $true; Repair = $true }).Count -gt 0)
Check '-Diagnostic + -PrinterName' (@(Test-ParameterCombination -Bound @{ Diagnostic = $true; PrinterName = 'x' }).Count -gt 0)
Check '-JobId + -PrinterName' (@(Test-ParameterCombination -Bound @{ JobId = 1; PrinterName = 'x' }).Count -gt 0)
Check '-Update + -ComputerName' (@(Test-ParameterCombination -Bound @{ Update = $true; ComputerName = 'a' }).Count -gt 0)
Check '-SaveConfig without -ConfigPath' (@(Test-ParameterCombination -Bound @{ SaveConfig = $true }).Count -gt 0)
Check '-SelfElevate + -NoElevate' (@(Test-ParameterCombination -Bound @{ SelfElevate = $true; NoElevate = $true }).Count -gt 0)
Check '-AutoUpdate + -CheckUpdate' (@(Test-ParameterCombination -Bound @{ AutoUpdate = $true; CheckUpdate = $true }).Count -gt 0)
Check '-BackupPath without -Backup' (@(Test-ParameterCombination -Bound @{ BackupPath = 'x' }).Count -gt 0)
Check '-OlderThan <= -NewerThan' (@(Test-ParameterCombination -Bound @{ OlderThan = '1h'; NewerThan = '2h' }).Count -gt 0)
Check '-PurgeDriverStore needs -Force' (@(Test-ParameterCombination -Bound @{ Repair = $true; PurgeDriverStore = $true }).Count -gt 0)
Check 'valid combination passes' (@(Test-ParameterCombination -Bound @{ Force = $true; Backup = $true; OlderThan = '2h' }).Count -eq 0)
Check '-PurgeDriverStore + -Force passes' (@(Test-ParameterCombination -Bound @{ Repair = $true; PurgeDriverStore = $true; Force = $true }).Count -eq 0)

Section '16. Reporting and aggregation'
$r1 = New-ResultObject -ComputerName 'PC1' -Status 'Success'
$r1.JobsRemoved = 3; $r1.FilesDeleted = 6; $r1.BytesDeleted = 2048; $r1.ServiceIsRunning = $true
$r2 = New-ResultObject -ComputerName 'PC2' -Status 'Failed'
$r2.Errors.Add('WinRM unreachable') | Out-Null
$sum = New-RunSummary -Results @($r1, $r2) -Arguments @{ Mode = 'Cleanup'; DryRun = $false } -LogPath 'C:\x.log' -ConfigSource 'Defaults' -Duration ([TimeSpan]::FromSeconds(12))
Check 'Summary aggregates files' ($sum.FilesDeleted -eq 6)
Check 'Summary aggregates bytes' ($sum.BytesDeleted -eq 2048)
Check 'Mixed results -> PartialSuccess' ($sum.OverallStatus -eq 'PartialSuccess')
Check 'PartialSuccess -> exit code 1' ((Get-ExitCodeForSummary -Summary $sum) -eq 1)
$ok = New-RunSummary -Results @($r1) -Arguments @{ Mode = 'Cleanup'; DryRun = $false } -Duration ([TimeSpan]::FromSeconds(1))
Check 'All success -> exit code 0' ((Get-ExitCodeForSummary -Summary $ok) -eq 0)
$dr = New-ResultObject -ComputerName 'PC3' -Status 'DryRun'
$dsum = New-RunSummary -Results @($dr) -Arguments @{ Mode = 'Cleanup'; DryRun = $true } -Duration ([TimeSpan]::FromSeconds(1))
Check 'Dry run -> exit code 0' ((Get-ExitCodeForSummary -Summary $dsum) -eq 0)
$dg = New-ResultObject -ComputerName 'PC4' -Status 'Diagnostic'
$dgsum = New-RunSummary -Results @($dg) -Arguments @{ Mode = 'Diagnostic'; DryRun = $false } -Duration ([TimeSpan]::FromSeconds(1))
Check 'Diagnostic -> exit code 11' ((Get-ExitCodeForSummary -Summary $dgsum) -eq 11)
$json = ConvertTo-JsonReport -Summary $sum
Check 'JSON report is valid JSON' ($null -ne ($json | ConvertFrom-Json)) "$($json.Length) chars"
Check 'JSON report carries schemaVersion' ($json -match '"schemaVersion":\s*"4\.0"')
$csv = ConvertTo-CsvReport -Summary $sum
Check 'CSV report has a header row' ($csv -match 'ComputerName')
Check 'CSV report has one row per target' ((@($csv -split "`n").Count) -ge 3)
$html = ConvertTo-HtmlReport -Summary $sum
Check 'HTML report is self-contained' (($html -match '<!DOCTYPE html>') -and ($html -notmatch 'src="http') -and ($html -notmatch 'href="http[^"]*\.(css|js)'))
Check 'HTML report contains both computers' (($html -match 'PC1') -and ($html -match 'PC2'))
$txt = ConvertTo-TextReport -Summary $sum
Check 'Text report lists both computers' (($txt -match 'PC1') -and ($txt -match 'PC2'))

Section '17. HTML escaping (injection defence)'
$inj = New-ResultObject -ComputerName '<script>alert(1)</script>' -Status 'Failed'
$inj.Errors.Add('<img src=x onerror=alert(1)>') | Out-Null
$isum = New-RunSummary -Results @($inj) -Arguments @{ Mode = 'Cleanup'; DryRun = $false } -Duration ([TimeSpan]::FromSeconds(1))
$ihtml = ConvertTo-HtmlReport -Summary $isum
Check 'Script tags are escaped' ($ihtml -notmatch '<script>alert')
Check 'Injected markup is escaped, not emitted' (($ihtml -notmatch '<img src=x') -and ($ihtml -match '&lt;img'))

Section '18. Platform handling'
$isWin = Test-RunningOnWindows
Check "Detects platform (IsWindows=$isWin)" $true
if (-not $isWin) {
    Check 'Spooler reported unavailable off Windows' (-not (Test-PrintSpoolerAvailable))
    Check 'Not an administrator off Windows' (-not (Test-RunningAsAdministrator))
    $null = Clear-PrintSpoolerQueue -Force -NoColor -Quiet 2>$null
    Check 'Cleanup exits 10 (UnsupportedPlatform)' ($Script:RuntimeState.ExitCode -eq 10) "exit=$($Script:RuntimeState.ExitCode)"
}
else {
    Check 'Spooler detected on Windows' (Test-PrintSpoolerAvailable)
}

Section '19. Logging'
$logTmp = Join-Path ([IO.Path]::GetTempPath()) ("psq-log-{0}.log" -f [Guid]::NewGuid().ToString('N'))
try {
    $resolved = Resolve-LogTarget -ExplicitPath $logTmp
    Check 'Explicit log path is writable' $resolved.Writable $resolved.Source
    $saved = $Script:RuntimeState.LogPath
    $Script:RuntimeState.LogPath = $logTmp
    Write-LogInfo 'test message one' -Component 'Unit' -Target 'PC1'
    Write-LogError 'test message two' -Component 'Unit'
    $content = @(Get-Content -LiteralPath $logTmp)
    Check 'Log records written' ($content.Count -eq 2) "$($content.Count) lines"
    Check 'No blank probe line' ($content[0].Trim().Length -gt 0)
    Check 'Pipe-delimited UTC format' ($content[0] -match '^\d{4}-\d{2}-\d{2}T[\d:\.]+Z\|\w{1,8}\s*\|Unit\|PC1\|test message one$') $content[0]
    $Script:RuntimeState.JsonLog = $true
    Write-LogInfo 'json line' -Component 'Unit'
    $content2 = @(Get-Content -LiteralPath $logTmp)
    Check 'JSON log line is valid JSON' ($null -ne ($content2[-1] | ConvertFrom-Json))
    $Script:RuntimeState.JsonLog = $false
    $Script:RuntimeState.MinLogLevelIndex = 4
    Write-LogInfo 'should be filtered out' -Component 'Unit'
    $content3 = @(Get-Content -LiteralPath $logTmp)
    Check 'Log level threshold filters records' ($content3.Count -eq $content2.Count)
    $Script:RuntimeState.MinLogLevelIndex = 2
    $Script:RuntimeState.LogPath = $saved
    Check 'Ring buffer captured entries' ((Get-LogBufferSnapshot).Count -gt 0)
}
finally { Remove-Item -LiteralPath $logTmp -Force -ErrorAction SilentlyContinue }

Section '20. Log rotation and retention'
$rotDir = Join-Path ([IO.Path]::GetTempPath()) ("psq-rot-{0}" -f [Guid]::NewGuid().ToString('N'))
try {
    $null = New-Item -ItemType Directory -Path $rotDir -Force
    $logF = Join-Path $rotDir 'PrintSpoolerCleanup.log'
    Set-Content -LiteralPath $logF -Value ('x' * 200000) -Encoding UTF8
    Check 'Rotation triggers over the size limit' (Invoke-LogRotation -Path $logF -MaxSizeMB 0.1 -MaxFileCount 3)
    Check 'Rotated generation created' (Test-Path -LiteralPath "$logF.1")
    Check 'Under-limit log is not rotated' (-not (Invoke-LogRotation -Path $logF -MaxSizeMB 50 -MaxFileCount 3))
    # Rotation renames the active file, so recreate it to test retention.
    Set-Content -LiteralPath $logF -Value 'fresh active log' -Encoding UTF8
    foreach ($n in @('PrintSpoolerCleanup.log.2', 'PrintSpoolerCleanup_20200101-000000.log')) {
        $old = Join-Path $rotDir $n
        Set-Content -LiteralPath $old -Value 'old' -Encoding UTF8
        (Get-Item -LiteralPath $old).LastWriteTime = (Get-Date).AddDays(-400)
    }
    $unrelated = Join-Path $rotDir 'UnrelatedNotes.txt'
    Set-Content -LiteralPath $unrelated -Value 'keep me' -Encoding UTF8
    (Get-Item -LiteralPath $unrelated).LastWriteTime = (Get-Date).AddDays(-400)
    $pruned = Invoke-LogRetention -Path $logF -RetentionDays 30
    Check 'Retention prunes old logs' ($pruned -ge 2) "pruned $pruned"
    Check 'Retention never touches the active log' (Test-Path -LiteralPath $logF)
    Check 'Retention ignores unrelated files' (Test-Path -LiteralPath $unrelated)
    Check 'Retention disabled at 0 days' ((Invoke-LogRetention -Path $logF -RetentionDays 0) -eq 0)
}
finally { Remove-Item -LiteralPath $rotDir -Recurse -Force -ErrorAction SilentlyContinue }

Section '21. Backup and restore round trip'
$bkDir = Join-Path ([IO.Path]::GetTempPath()) ("psq-bk-{0}" -f [Guid]::NewGuid().ToString('N'))
try {
    $null = New-Item -ItemType Directory -Path $bkDir -Force
    'SPLDATA' | Set-Content -LiteralPath (Join-Path $bkDir '00001.SPL') -Encoding ASCII
    'SHDDATA' | Set-Content -LiteralPath (Join-Path $bkDir 'F00001.SHD') -Encoding ASCII
    $manifest = [ordered]@{
        Tool = 'Clear-PrintSpoolerQueue'; ToolVersion = '4.0.0'
        CreatedUtc = [DateTime]::UtcNow.ToString('o'); ComputerName = 'TESTPC'
        SpoolDirectory = 'C:\Windows\System32\spool\PRINTERS'; FileCount = 2
        Files = @(@{ Name = '00001.SPL'; Size = 7 }, @{ Name = 'F00001.SHD'; Size = 7 })
        Jobs = @(@{ JobId = 1; Printer = 'Test'; Document = 'x.pdf'; Owner = 'u'; Size = 7 })
        RestoreNote = 'test'
    }
    ($manifest | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath (Join-Path $bkDir 'backup-manifest.json') -Encoding UTF8
    $zip = Join-Path $bkDir 'spool-backup_TESTPC_20260101-000000.zip'
    Compress-Archive -Path (Join-Path $bkDir '*') -DestinationPath $zip -CompressionLevel Optimal -Force
    $info = Get-SpoolBackupInfo -Path $zip
    Check 'Backup manifest is readable' $info.Valid $info.Error
    Check 'Backup records the computer' ($info.ComputerName -eq 'TESTPC')
    Check 'Backup records the file count' ($info.FileCount -eq 2)
    Check 'Backup SHA-256 is computed' ($info.Hash.Length -eq 64)
    $dest = Join-Path $bkDir 'restored'
    $res = Restore-SpoolerBackupInternal -Path $zip -DestinationDirectory $dest
    Check 'Restore extracts the files' ($res.FilesRestored -ge 3) "$($res.FilesRestored) files"
    Check 'Restore lands in the destination' (Test-Path -LiteralPath (Join-Path $dest '00001.SPL'))
    Check 'Restore states the honest limitation' ($res.Note -match 'cannot be re-queued')
    $thrown = $false
    try { Restore-SpoolerBackupInternal -Path $zip -DestinationDirectory (Join-Path $bkDir 'r2') -ExpectedHash ('0' * 64) | Out-Null } catch { $thrown = $true }
    Check 'Restore rejects a hash mismatch' $thrown
    $okHash = $false
    try { $null = Restore-SpoolerBackupInternal -Path $zip -DestinationDirectory (Join-Path $bkDir 'r3') -ExpectedHash $info.Hash; $okHash = $true } catch { }
    Check 'Restore accepts the correct hash' $okHash
    Check 'Backup inventory lists the bundle' (@(Get-SpoolBackupInventory -BackupPath $bkDir).Count -eq 1)
    $badInfo = Get-SpoolBackupInfo -Path (Join-Path $bkDir 'does-not-exist.zip')
    Check 'Missing archive reports an error' (-not $badInfo.Valid -and $badInfo.Error)
}
catch { Check 'Backup round trip' $false $_.Exception.Message }
finally { Remove-Item -LiteralPath $bkDir -Recurse -Force -ErrorAction SilentlyContinue }

Section '22. Export round trip'
$exDir = Join-Path ([IO.Path]::GetTempPath()) ("psq-ex-{0}" -f [Guid]::NewGuid().ToString('N'))
try {
    $null = New-Item -ItemType Directory -Path $exDir -Force
    $files = @(Export-RunReport -Summary $sum -Format @('Json','Csv','Html','Text') -Path $exDir)
    Check 'All four formats written' ($files.Count -eq 4) (($files | ForEach-Object { Split-Path -Leaf $_ }) -join ', ')
    foreach ($f in $files) { Check "$(Split-Path -Leaf $f) is non-empty" ((Get-Item -LiteralPath $f).Length -gt 50) }
    $jsonFile = @($files | Where-Object { $_ -like '*.json' })[0]
    Check 'Exported JSON reparses' ($null -ne (Get-Content -LiteralPath $jsonFile -Raw | ConvertFrom-Json))
}
finally { Remove-Item -LiteralPath $exDir -Recurse -Force -ErrorAction SilentlyContinue }

Section '23. Self-elevation argument reconstruction'
$bound = @{
    Force = [switch]::Present; Backup = [switch]::Present; OlderThan = '2h'
    ComputerName = @('SRV 01','SRV02')
    Credential = (New-Object System.Management.Automation.PSCredential('u', (ConvertTo-SecureString 'p' -AsPlainText -Force)))
    PreFlightProbe = $true; ThrottleLimit = 8
}
$child = @(New-ChildArgumentList -BoundParameters $bound)
$joined = $child -join ' '
Check 'Switches forwarded bare' ($child -contains '-Force')
Check 'Scalar values forwarded' (($child -contains '-OlderThan') -and ($child -contains '2h'))
Check 'Values with spaces are quoted' ($child -contains "'SRV 01',SRV02")
Check 'Booleans forwarded as -Name:$true' ($child -contains '-PreFlightProbe:$true') $joined
Check 'Credentials are never forwarded' ($joined -notmatch 'Credential')
Check 'Numbers forwarded' ($child -contains '8')

Section '24. Selection description'
Check 'No filters says ALL' ((Format-SelectionDescription -Arguments @{}) -match 'ALL spool files')
Check 'Printer filter described' ((Format-SelectionDescription -Arguments @{ PrinterName = @('HP*') }) -match "printers like 'HP\*'")
Check 'Age filter described' ((Format-SelectionDescription -Arguments @{ OlderThan = '2h' }) -match 'older than 2h')
Check 'Combined filters described' ((Format-SelectionDescription -Arguments @{ PrinterName = @('HP*'); OlderThan = '90m' }) -match 'jobs matching')
Check 'Job ids described' ((Format-SelectionDescription -Arguments @{ JobId = @(7, 8) }) -match 'job id\(s\) 7, 8')

Section '25. Local/remote target detection'
foreach ($n in @('localhost','.','127.0.0.1','::1','(local)')) { Check "'$n' is local" (Test-IsLocalTarget -ComputerName $n) }
Check "'SRV01' is remote" (-not (Test-IsLocalTarget -ComputerName 'SRV01'))
Check "'srv01.example.com' is remote" (-not (Test-IsLocalTarget -ComputerName 'srv01.example.com'))
$hn = [System.Net.Dns]::GetHostName()
Check "FQDN of this host is local" (Test-IsLocalTarget -ComputerName "$hn.example.com")

Section '26. Update discovery'
if ($SkipNetwork) {
    Check 'Live discovery skipped (-SkipNetwork)' $true
}
else {
    try {
        $info = Get-UpdateInfoCore -Owner 'paulmann' -Repo 'Print-Spooler-Queue-Cleanup' -CurrentVersion '0.0.1'
        Check 'Discovery returns a result' ($null -ne $info)
        Write-Host ("        source={0} latest={1} errors={2}" -f $info.Source, $info.LatestVersion, ($info.Errors -join ' | ')) -ForegroundColor DarkGray
        Check 'A version was discovered' (-not [string]::IsNullOrWhiteSpace($info.LatestVersion)) "v$($info.LatestVersion) via $($info.Source)"
        if ($info.LatestVersion) {
            Check 'Far-future version reports no update' (-not (Get-UpdateInfoCore -Owner 'paulmann' -Repo 'Print-Spooler-Queue-Cleanup' -CurrentVersion '99.0.0').UpdateAvailable)
            Check 'v1.0.0 reports an update' (Get-UpdateInfoCore -Owner 'paulmann' -Repo 'Print-Spooler-Queue-Cleanup' -CurrentVersion '1.0.0').UpdateAvailable
        }
    } catch { Check 'Live discovery' $false $_.Exception.Message }
}
Check 'No state -> cooldown inactive' (-not (Test-UpdateCooldownActive -State @{} -CooldownHours 24))
Check 'Fresh check -> cooldown active' (Test-UpdateCooldownActive -State @{ LastCheckUtc = [DateTime]::UtcNow.ToString('o') } -CooldownHours 24)
Check 'Old check -> cooldown inactive' (-not (Test-UpdateCooldownActive -State @{ LastCheckUtc = [DateTime]::UtcNow.AddHours(-30).ToString('o') } -CooldownHours 24))
Check 'Cooldown 0 disables the wait' (-not (Test-UpdateCooldownActive -State @{ LastCheckUtc = [DateTime]::UtcNow.ToString('o') } -CooldownHours 0))
$stTmp = Join-Path ([IO.Path]::GetTempPath()) ("psq-state-{0}.json" -f [Guid]::NewGuid().ToString('N'))
try {
    $null = Write-UpdateState -Path $stTmp -State @{ LastCheckUtc = [DateTime]::UtcNow.ToString('o'); LastKnownVersion = '9.9.9' }
    Check 'State round-trips through JSON' ((Read-UpdateState -Path $stTmp)['LastKnownVersion'] -eq '9.9.9')
    Check 'Missing state file is tolerated' ($null -ne (Read-UpdateState -Path (Join-Path ([IO.Path]::GetTempPath()) 'no-such-file.json')))
}
finally { Remove-Item -LiteralPath $stTmp -Force -ErrorAction SilentlyContinue }

Section '27. Version stamp consistency across artifacts'
$distDir = Join-Path $Root 'dist'
$psVer = (Get-Content -LiteralPath (Join-Path $distDir 'Clear-PrintSpoolerQueue.ps1') -Raw)
$batText = (Get-Content -LiteralPath (Join-Path $distDir 'Clear-PrintSpoolerQueue.bat') -Raw)
$psdText = (Get-Content -LiteralPath (Join-Path $distDir 'PrintSpoolerCleanup/PrintSpoolerCleanup.psd1') -Raw)
$verMatch = [regex]::Match($psVer, "ScriptVersion\s+=\s+'([^']+)'")
$expectedVer = if ($verMatch.Success) { $verMatch.Groups[1].Value } else { '' }
$expectedPattern = [regex]::Escape($expectedVer)
# The expected version is read from the artifact under test. Hardcoding it here
# meant every release had to edit the test, and a stale literal would have kept
# passing while the artifacts disagreed with each other.
Check '.ps1 carries a version stamp' ($expectedVer -ne '') $expectedVer
Check '.bat carries the same version' ($batText -match ("SCRIPT_VERSION=" + $expectedPattern))
Check '.psd1 carries the same version' ($psdText -match ("ModuleVersion\s+=\s+'" + $expectedPattern + "'"))
Check '.ps1 has no @@BUILD_STAMP@@ left' ($psVer -notmatch '@@BUILD_STAMP@@')
Check '.ps1 declares the GPL license' ($psVer -match 'GPL-3\.0-or-later')

Write-Host ''
Write-Host ('=' * 78) -ForegroundColor Cyan
Write-Host ("  {0} passed, {1} failed" -f $script:passes, $script:failures) -ForegroundColor $(if ($script:failures -eq 0) { 'Green' } else { 'Red' })
Write-Host ('=' * 78) -ForegroundColor Cyan
exit $(if ($script:failures -eq 0) { 0 } else { 1 })

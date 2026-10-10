# Unit tests: parsers, formatting, logging, rotation, checksums, exit codes, reports
Start-TestSuite 'Time-span parsing'

$spanCases = @(
    @{ In = '30s';   Prop = 'TotalSeconds';  Val = 30 }
    @{ In = '90s';   Prop = 'TotalSeconds';  Val = 90 }
    @{ In = '15m';   Prop = 'TotalMinutes';  Val = 15 }
    @{ In = '45min'; Prop = 'TotalMinutes';  Val = 45 }
    @{ In = '2h';    Prop = 'TotalHours';    Val = 2 }
    @{ In = '1hr';   Prop = 'TotalHours';    Val = 1 }
    @{ In = '7d';    Prop = 'TotalDays';     Val = 7 }
    @{ In = '1w';    Prop = 'TotalDays';     Val = 7 }
    @{ In = '500ms'; Prop = 'TotalMilliseconds'; Val = 500 }
    @{ In = '1.5h';  Prop = 'TotalMinutes';  Val = 90 }
    @{ In = '30';    Prop = 'TotalSeconds';  Val = 30 }
    @{ In = '45';    Prop = 'TotalSeconds';  Val = 45 }
)
foreach ($case in $spanCases) {
    $c = $case
    Invoke-TestCase -Name "'$($c.In)' -> $($c.Val) $($c.Prop)" -Body {
        $ts = ConvertFrom-AnyTimeSpan -Value $c.In
        Assert-Equal $c.Val $ts.($c.Prop)
    }
}

Invoke-TestCase -Name 'A standard TimeSpan string parses' -Body {
    Assert-Equal 90 (ConvertFrom-AnyTimeSpan -Value '01:30:00').TotalMinutes
    Assert-Equal 25 (ConvertFrom-AnyTimeSpan -Value '1.01:00:00').TotalHours
}

Invoke-TestCase -Name 'A TimeSpan object passes through unchanged' -Body {
    $ts = [TimeSpan]::FromMinutes(5)
    Assert-Equal 5 (ConvertFrom-AnyTimeSpan -Value $ts).TotalMinutes
}

Invoke-TestCase -Name 'Numeric input is interpreted as seconds' -Body {
    Assert-Equal 45 (ConvertFrom-AnyTimeSpan -Value 45).TotalSeconds
    Assert-Equal 90 (ConvertFrom-AnyTimeSpan -Value 90.0).TotalSeconds
}

Invoke-TestCase -Name 'Null and empty yield zero' -Body {
    Assert-Equal 0 (ConvertFrom-AnyTimeSpan -Value $null).TotalSeconds
    Assert-Equal 0 (ConvertFrom-AnyTimeSpan -Value '').TotalSeconds
    Assert-Equal 0 (ConvertFrom-AnyTimeSpan -Value '   ').TotalSeconds
}

Invoke-TestCase -Name 'Garbage throws with an actionable message' -Body {
    $msg = $null
    try { ConvertFrom-AnyTimeSpan -Value 'nonsense' | Out-Null } catch { $msg = $_.Exception.Message }
    Assert-NotNull $msg
    Assert-Match 'Cannot interpret' $msg
    Assert-Match '2h' $msg 'the message must show the accepted formats'
}

Invoke-TestCase -Name 'Case and surrounding whitespace are tolerated' -Body {
    Assert-Equal 2 (ConvertFrom-AnyTimeSpan -Value '  2H  ').TotalHours
    Assert-Equal 7 (ConvertFrom-AnyTimeSpan -Value '7D').TotalDays
}

Start-TestSuite 'Byte-size parsing and formatting'

Invoke-TestCase -Name 'Human sizes parse to exact byte counts' -Body {
    Assert-Equal 512 (Format-FileSizeArgument -Value '512')
    Assert-Equal 1024 (Format-FileSizeArgument -Value '1KB')
    Assert-Equal 10485760 (Format-FileSizeArgument -Value '10MB')
    Assert-Equal 1610612736 (Format-FileSizeArgument -Value '1.5GB')
    Assert-Equal 1099511627776 (Format-FileSizeArgument -Value '1TB')
    Assert-Equal 2048 (Format-FileSizeArgument -Value '2k')
}

Invoke-TestCase -Name 'Null, empty and garbage yield null (not zero)' -Body {
    Assert-Null (Format-FileSizeArgument -Value $null)
    Assert-Null (Format-FileSizeArgument -Value '')
    Assert-Null (Format-FileSizeArgument -Value 'banana')
    Assert-Null (Format-FileSizeArgument -Value 'MB')
}

Invoke-TestCase -Name 'Formatting is locale-independent' -Body {
    Assert-Equal '0 B' (Format-ByteSize 0)
    Assert-Equal '512 B' (Format-ByteSize 512)
    Assert-Equal '1.50 KB' (Format-ByteSize 1536)
    Assert-Equal '1.00 MB' (Format-ByteSize 1MB)
    Assert-Equal '2.50 GB' (Format-ByteSize 2.5GB)
    Assert-Equal '0 B' (Format-ByteSize $null)
}

Invoke-TestCase -Name 'Negative sizes keep their sign' -Body {
    Assert-Match '^-' (Format-ByteSize -1536)
}

Invoke-TestCase -Name 'Duration formatting is stable' -Body {
    Assert-Equal '00:00:00.000' (Format-Duration -Duration ([TimeSpan]::Zero))
    Assert-Equal '01:02:03.004' (Format-Duration -Duration ([TimeSpan]::new(0, 1, 2, 3, 4)))
    Assert-Match '^2d ' (Format-Duration -Duration ([TimeSpan]::FromDays(2.5)))
    Assert-Equal '00:00:00.000' (Format-Duration -Duration $null)
}

Start-TestSuite 'File-name safety'

Invoke-TestCase -Name 'Windows-illegal characters are neutralised on any host' -Body {
    # GetInvalidFileNameChars() reflects the HOST platform, so the function must
    # carry the Windows set explicitly - the product only ever runs on Windows.
    $out = ConvertTo-SafeFileName -Value 'a/b:c*d?"<>|'
    Assert-NotMatch '[\\/:*?"<>|]' $out
}

Invoke-TestCase -Name 'Empty and whitespace input yields a usable name' -Body {
    Assert-Equal 'unnamed' (ConvertTo-SafeFileName -Value '')
    Assert-Equal 'unnamed' (ConvertTo-SafeFileName -Value '   ')
    Assert-Equal 'unnamed' (ConvertTo-SafeFileName -Value $null)
    Assert-Equal 'unnamed' (ConvertTo-SafeFileName -Value '///')
}

Invoke-TestCase -Name 'MaxLength is honoured' -Body {
    Assert-Equal 20 (ConvertTo-SafeFileName -Value ('x' * 300) -MaxLength 20).Length
}

Invoke-TestCase -Name 'Spaces become the replacement character' -Body {
    Assert-NotMatch ' ' (ConvertTo-SafeFileName -Value 'SRV 01 print queue')
}

Invoke-TestCase -Name 'Timestamps are culture-invariant' -Body {
    $stamp = Get-SafeFileNameStamp -When ([DateTime]::new(2026, 3, 4, 5, 6, 7))
    Assert-Equal '20260304-050607' $stamp
}

Start-TestSuite 'Checksum handling'

Invoke-TestCase -Name 'GNU coreutils format parses' -Body {
    $txt = "abc123def456abc123def456abc123def456abc123def456abc123def456abc1  file.ps1"
    $m = ConvertFrom-ChecksumText -Text $txt
    Assert-True $m.Contains('file.ps1')
}

Invoke-TestCase -Name 'The binary-mode asterisk marker is stripped' -Body {
    $txt = "abc123def456abc123def456abc123def456abc123def456abc123def456abc1 *file.zip"
    Assert-True (ConvertFrom-ChecksumText -Text $txt).Contains('file.zip')
}

Invoke-TestCase -Name 'The reversed "name: hash" format parses' -Body {
    $txt = "file.ps1: abc123def456abc123def456abc123def456abc123def456abc123def456abc1"
    Assert-True (ConvertFrom-ChecksumText -Text $txt).Contains('file.ps1')
}

Invoke-TestCase -Name 'Comments, blank lines and short hashes are ignored' -Body {
    $txt = "# a comment`r`n`r`ntooshort  file.ps1`r`nabc123def456abc123def456abc123def456abc123def456abc123def456abc1  good.ps1"
    $m = ConvertFrom-ChecksumText -Text $txt
    Assert-Equal 1 $m.Count
    Assert-True $m.Contains('good.ps1')
}

Invoke-TestCase -Name 'Empty and null input yields an empty map' -Body {
    Assert-Equal 0 (ConvertFrom-ChecksumText -Text '').Count
    Assert-Equal 0 (ConvertFrom-ChecksumText -Text $null).Count
}

Invoke-TestCase -Name 'Hash comparison is case-insensitive and format-tolerant' -Body {
    Assert-True (Test-HashMatch -Expected 'ABC123' -Actual 'abc123')
    Assert-True (Test-HashMatch -Expected 'abc123' -Actual '  ABC123  ')
    Assert-True (Test-HashMatch -Expected 'AB:CD:12' -Actual 'abcd12')
}

Invoke-TestCase -Name 'Hash comparison rejects mismatches and empties' -Body {
    Assert-False (Test-HashMatch -Expected 'abc123' -Actual 'abc124')
    Assert-False (Test-HashMatch -Expected 'abc123' -Actual 'abc12')
    Assert-False (Test-HashMatch -Expected '' -Actual 'abc123')
    Assert-False (Test-HashMatch -Expected 'abc123' -Actual '')
    Assert-False (Test-HashMatch -Expected $null -Actual $null)
}

Invoke-TestCase -Name 'Generation and parsing round-trip' -Body {
    $f = Join-Path ([IO.Path]::GetTempPath()) ("psq-hash-{0}.txt" -f [Guid]::NewGuid().ToString('N'))
    try {
        Set-Content -LiteralPath $f -Value 'payload' -Encoding UTF8
        $text = ConvertTo-ChecksumText -FilePaths @($f)
        $m = ConvertFrom-ChecksumText -Text $text
        Assert-True $m.Contains((Split-Path -Leaf $f))
        Assert-Equal (Get-FileHashSafe -Path $f).ToLowerInvariant() $m[(Split-Path -Leaf $f)]
    }
    finally { Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue }
}

Invoke-TestCase -Name 'Get-FileHashSafe returns null instead of throwing' -Body {
    Assert-Null (Get-FileHashSafe -Path (Join-Path ([IO.Path]::GetTempPath()) 'definitely-missing.bin'))
}

Start-TestSuite 'Exit-code contract'

Invoke-TestCase -Name 'Codes are unique' -Body {
    $codes = @($Script:ExitCode.Values)
    Assert-Equal $codes.Count @($codes | Select-Object -Unique).Count
}

Invoke-TestCase -Name 'Success is 0 and the documented values are stable' -Body {
    Assert-Equal 0 $Script:ExitCode.Success
    Assert-Equal 1 $Script:ExitCode.PartialSuccess
    Assert-Equal 2 $Script:ExitCode.Failed
    Assert-Equal 3 $Script:ExitCode.BadArguments
    Assert-Equal 4 $Script:ExitCode.Cancelled
    Assert-Equal 5 $Script:ExitCode.NotAdministrator
    Assert-Equal 7 $Script:ExitCode.UpdateAvailable
    Assert-Equal 11 $Script:ExitCode.DiagnosticsOnly
    Assert-Equal 12 $Script:ExitCode.ServiceFailure
    Assert-Equal 130 $Script:ExitCode.Aborted
}

Invoke-TestCase -Name 'Every code has a human-readable meaning' -Body {
    foreach ($row in (Get-ExitCodeMeaning)) {
        Assert-True (-not [string]::IsNullOrWhiteSpace($row.Description)) "code $($row.Code) has no description"
        Assert-True (-not [string]::IsNullOrWhiteSpace($row.Name)) "code $($row.Code) has no name"
    }
}

Invoke-TestCase -Name 'Get-ExitCodeMeaning looks up a single code' -Body {
    Assert-Equal 'ServiceFailure' (Get-ExitCodeMeaning -Code 12).Name
    Assert-Null (Get-ExitCodeMeaning -Code 999)
}

Invoke-TestCase -Name 'Summary status maps to the documented exit code' -Body {
    $mk = {
        param($status)
        $r = New-ResultObject -ComputerName 'PC' -Status $status
        New-RunSummary -Results @($r) -Arguments @{ Mode = 'Cleanup'; DryRun = $false } -Duration ([TimeSpan]::FromSeconds(1))
    }
    Assert-Equal 0 (Get-ExitCodeForSummary -Summary (& $mk 'Success'))
    Assert-Equal 0 (Get-ExitCodeForSummary -Summary (& $mk 'NothingToDo'))
    Assert-Equal 0 (Get-ExitCodeForSummary -Summary (& $mk 'DryRun'))
    Assert-Equal 11 (Get-ExitCodeForSummary -Summary (& $mk 'Diagnostic'))
    Assert-Equal 1 (Get-ExitCodeForSummary -Summary (& $mk 'SuccessWithWarnings'))
    Assert-Equal 2 (Get-ExitCodeForSummary -Summary (& $mk 'Failed'))
    Assert-Equal 2 (Get-ExitCodeForSummary -Summary (& $mk 'ServiceFailure'))
}

Start-TestSuite 'Result aggregation'

Invoke-TestCase -Name 'Counters are summed across targets' -Body {
    $a = New-ResultObject -ComputerName 'A' -Status 'Success'
    $a.JobsFound = 3; $a.JobsRemoved = 3; $a.FilesDeleted = 6; $a.BytesDeleted = 1000
    $b = New-ResultObject -ComputerName 'B' -Status 'Success'
    $b.JobsFound = 2; $b.JobsRemoved = 2; $b.FilesDeleted = 4; $b.BytesDeleted = 500
    $s = New-RunSummary -Results @($a, $b) -Arguments @{ Mode = 'Cleanup'; DryRun = $false } -Duration ([TimeSpan]::FromSeconds(1))
    Assert-Equal 5 $s.JobsFound
    Assert-Equal 5 $s.JobsRemoved
    Assert-Equal 10 $s.FilesDeleted
    Assert-Equal 1500 $s.BytesDeleted
    Assert-Equal 2 $s.TargetsSucceeded
}

Invoke-TestCase -Name 'One failure among successes is PartialSuccess' -Body {
    $a = New-ResultObject -ComputerName 'A' -Status 'Success'
    $b = New-ResultObject -ComputerName 'B' -Status 'Failed'
    $s = New-RunSummary -Results @($a, $b) -Arguments @{ Mode = 'Cleanup'; DryRun = $false } -Duration ([TimeSpan]::FromSeconds(1))
    Assert-Equal 'PartialSuccess' $s.OverallStatus
}

Invoke-TestCase -Name 'All failures is Failed' -Body {
    $a = New-ResultObject -ComputerName 'A' -Status 'Failed'
    $b = New-ResultObject -ComputerName 'B' -Status 'ServiceFailure'
    $s = New-RunSummary -Results @($a, $b) -Arguments @{ Mode = 'Cleanup'; DryRun = $false } -Duration ([TimeSpan]::FromSeconds(1))
    Assert-Equal 'Failed' $s.OverallStatus
}

Invoke-TestCase -Name 'No results at all is Failed, not Success' -Body {
    $s = New-RunSummary -Results @() -Arguments @{ Mode = 'Cleanup'; DryRun = $false } -Duration ([TimeSpan]::FromSeconds(1))
    Assert-Equal 'Failed' $s.OverallStatus
}

Invoke-TestCase -Name 'Warnings and errors are counted from every target' -Body {
    $a = New-ResultObject -ComputerName 'A' -Status 'SuccessWithWarnings'
    $a.Warnings.Add('w1') | Out-Null; $a.Warnings.Add('w2') | Out-Null
    $b = New-ResultObject -ComputerName 'B' -Status 'Failed'
    $b.Errors.Add('e1') | Out-Null
    $s = New-RunSummary -Results @($a, $b) -Arguments @{ Mode = 'Cleanup'; DryRun = $false } -Duration ([TimeSpan]::FromSeconds(1))
    Assert-Equal 2 $s.Warnings
    Assert-Equal 1 $s.Errors
}

Invoke-TestCase -Name 'Backup paths are collected for the report' -Body {
    $a = New-ResultObject -ComputerName 'A' -Status 'Success'
    $a.BackupPath = 'C:\backups\a.zip'
    $s = New-RunSummary -Results @($a) -Arguments @{ Mode = 'Cleanup'; DryRun = $false } -Duration ([TimeSpan]::FromSeconds(1))
    Assert-Equal 1 @($s.BackupsCreated).Count
}

Invoke-TestCase -Name 'The summary records the host environment for auditability' -Body {
    $s = New-RunSummary -Results @() -Arguments @{ Mode = 'Cleanup'; DryRun = $false } -Duration ([TimeSpan]::FromSeconds(1))
    Assert-Equal $Script:ScriptVersion $s.ToolVersion
    Assert-True (-not [string]::IsNullOrWhiteSpace($s.PowerShellVersion))
    Assert-True (-not [string]::IsNullOrWhiteSpace($s.StartedUtc))
}

Start-TestSuite 'Report serialisation'

$sumForReports = {
    $a = New-ResultObject -ComputerName 'PC-ONE' -Status 'Success'
    $a.JobsRemoved = 2; $a.FilesDeleted = 4; $a.BytesDeleted = 4096; $a.ServiceIsRunning = $true
    $a.SpoolDirectory = 'C:\Windows\System32\spool\PRINTERS'
    $a.BackupPath = 'C:\backups\a.zip'; $a.BackupHash = ('A' * 64)
    $a.Jobs = @([pscustomobject]@{ JobId = 7; PrinterName = 'HP <test>'; DocumentName = 'a&b.pdf'; UserName = 'CORP\jdoe'; SizeBytes = 1024; TotalPages = 3; Status = 'Printing'; SubmittedUtc = [DateTime]::UtcNow.ToString('o'); Source = 'WMI' })
    $b = New-ResultObject -ComputerName 'PC-TWO' -Status 'Failed'
    $b.Errors.Add('WinRM unreachable') | Out-Null
    $b.Warnings.Add('probe timed out') | Out-Null
    New-RunSummary -Results @($a, $b) -Arguments @{ Mode = 'Cleanup'; DryRun = $false } -LogPath 'C:\logs\x.log' -ConfigSource 'Defaults' -Duration ([TimeSpan]::FromSeconds(9))
}
$reportSummary = & $sumForReports

Invoke-TestCase -Name 'JSON report is valid and carries the schema version' -Body {
    $json = ConvertTo-JsonReport -Summary $reportSummary
    $obj = $json | ConvertFrom-Json
    Assert-NotNull $obj
    Assert-Equal '4.0' $obj.schemaVersion
    Assert-Equal 'PC-ONE' $obj.summary.Results[0].ComputerName
}

Invoke-TestCase -Name 'JSON report survives hostile characters in job metadata' -Body {
    $json = ConvertTo-JsonReport -Summary $reportSummary
    Assert-NotThrows { $null = $json | ConvertFrom-Json }
}

Invoke-TestCase -Name 'CSV report has one header and one row per target' -Body {
    $csv = ConvertTo-CsvReport -Summary $reportSummary
    $lines = @($csv -split "`r?`n" | Where-Object { $_.Trim() })
    Assert-Equal 3 $lines.Count
    Assert-Match 'ComputerName' $lines[0]
    Assert-Match 'PC-ONE' $lines[1]
}

Invoke-TestCase -Name 'CSV report imports cleanly' -Body {
    $csv = ConvertTo-CsvReport -Summary $reportSummary
    $rows = @($csv | ConvertFrom-Csv)
    Assert-Equal 2 $rows.Count
    Assert-Equal 'PC-TWO' $rows[1].ComputerName
    Assert-Equal 'WinRM unreachable' $rows[1].FirstError
}

Invoke-TestCase -Name 'CSV of an empty run is empty, not malformed' -Body {
    $empty = New-RunSummary -Results @() -Arguments @{ Mode = 'Cleanup'; DryRun = $false } -Duration ([TimeSpan]::FromSeconds(1))
    Assert-Equal '' (ConvertTo-CsvReport -Summary $empty)
}

Invoke-TestCase -Name 'HTML report is self-contained' -Body {
    $html = ConvertTo-HtmlReport -Summary $reportSummary
    Assert-Match '<!DOCTYPE html>' $html
    Assert-NotMatch '<script' $html 'inline script is not acceptable in a report that gets attached to tickets'
    Assert-NotMatch 'src="http' $html
    Assert-NotMatch '<link[^>]+href="http' $html
    Assert-Match '<style>' $html
}

Invoke-TestCase -Name 'HTML report escapes injected markup' -Body {
    $inj = New-ResultObject -ComputerName '<script>alert(1)</script>' -Status 'Failed'
    $inj.Errors.Add('<img src=x onerror=alert(1)>') | Out-Null
    $inj.Jobs = @([pscustomobject]@{ JobId = 1; PrinterName = '"><b>'; DocumentName = '&'; UserName = "'"; SizeBytes = 1; TotalPages = 1; Status = 'x'; SubmittedUtc = $null; Source = 'WMI' })
    $s = New-RunSummary -Results @($inj) -Arguments @{ Mode = 'Cleanup'; DryRun = $false } -Duration ([TimeSpan]::FromSeconds(1))
    $html = ConvertTo-HtmlReport -Summary $s
    Assert-NotMatch '<script>alert' $html
    Assert-NotMatch '<img src=x' $html
    Assert-Match '&lt;script&gt;' $html
    Assert-Match '&lt;img' $html
}

Invoke-TestCase -Name 'Text report contains the audit essentials' -Body {
    $txt = ConvertTo-TextReport -Summary $reportSummary
    Assert-Match 'PC-ONE' $txt
    Assert-Match 'PC-TWO' $txt
    Assert-Match 'ERROR: WinRM unreachable' $txt
    Assert-Match 'C:\\Windows\\System32\\spool\\PRINTERS' $txt
    Assert-Match $Script:ScriptVersion $txt
}

Start-TestSuite 'Logging'

$logDir = Join-Path ([IO.Path]::GetTempPath()) ("psq-log-{0}" -f [Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $logDir -Force
$logFile = Join-Path $logDir 'unit.log'
$savedLogPath = $Script:RuntimeState.LogPath
$savedJson = $Script:RuntimeState.JsonLog
$Script:RuntimeState.LogPath = $logFile

Invoke-TestCase -Name 'Records use the documented pipe-delimited UTC schema' -Body {
    Set-Content -LiteralPath $logFile -Value '' -Encoding UTF8 -NoNewline
    Write-LogInfo 'hello world' -Component 'Unit' -Target 'PC1'
    $line = @(Get-Content -LiteralPath $logFile | Where-Object { $_.Trim() })[0]
    Assert-Match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z\|INFO    \|Unit\|PC1\|hello world$' $line
}

Invoke-TestCase -Name 'A missing target renders as a dash' -Body {
    Write-LogWarn 'no target here' -Component 'Unit'
    $lines = @(Get-Content -LiteralPath $logFile | Where-Object { $_.Trim() })
    Assert-Match '\|WARN    \|Unit\|-\|no target here$' $lines[-1]
}

Invoke-TestCase -Name 'Newlines in a message cannot forge extra records' -Body {
    Write-LogInfo "line one`r`nline two" -Component 'Unit'
    $lines = @(Get-Content -LiteralPath $logFile | Where-Object { $_.Trim() })
    Assert-Match 'line one line two$' $lines[-1]
}

Invoke-TestCase -Name 'The level threshold filters records' -Body {
    $Script:RuntimeState.MinLogLevelIndex = 4
    $before = @(Get-Content -LiteralPath $logFile).Count
    Write-LogInfo 'filtered out' -Component 'Unit'
    Write-LogError 'kept' -Component 'Unit'
    $after = @(Get-Content -LiteralPath $logFile).Count
    Assert-Equal ($before + 1) $after
    # Restored here rather than relying on a variable captured earlier: an
    # undefined name would silently assign $null, which makes every subsequent
    # `-ge` comparison false and disables file logging for the rest of the suite.
    $Script:RuntimeState.MinLogLevelIndex = [int]$Script:LogLevels['INFO']
}

Invoke-TestCase -Name 'JSON log mode emits one parseable object per line' -Body {
    # Preconditions are set explicitly rather than inherited: a test that depends
    # on the residue of an earlier test passes alone and fails in the suite, which
    # is the worst possible failure mode to debug.
    $jsonLog = Join-Path $logDir 'unit.json.log'
    Remove-Item -LiteralPath $jsonLog -Force -ErrorAction SilentlyContinue
    $Script:RuntimeState.LogPath = $jsonLog
    $Script:RuntimeState.JsonLog = $true
    $Script:RuntimeState.MinLogLevelIndex = [int]$Script:LogLevels['TRACE']
    $Script:RuntimeState.MachineReadable = $false
    $Script:RuntimeState.LogFailureWarned = $false

    Write-LogInfo 'structured' -Component 'Unit' -Target 'PC9'
    Write-LogWarn 'second record' -Component 'Unit'

    $Script:RuntimeState.JsonLog = $false
    $Script:RuntimeState.LogPath = $logFile
    $Script:RuntimeState.MinLogLevelIndex = [int]$Script:LogLevels['INFO']

    $lines = @(Get-Content -LiteralPath $jsonLog | Where-Object { $_.Trim() })
    Assert-Equal 2 $lines.Count 'one JSON object per record, nothing else'
    $obj = $lines[0] | ConvertFrom-Json
    Assert-Equal 'INFO' $obj.level
    Assert-Equal 'Unit' $obj.component
    Assert-Equal 'PC9' $obj.target
    Assert-Equal 'structured' $obj.msg
    # The RAW line must carry the documented UTC ISO-8601 stamp. Asserting on
    # $obj.ts is not enough: PowerShell 7's ConvertFrom-Json auto-converts an
    # ISO-8601 string into a [datetime], so the property no longer holds the
    # wire format that a log shipper would actually parse.
    Assert-Match '"ts":"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z"' $lines[0]
    Assert-Equal 'WARN' ($lines[1] | ConvertFrom-Json).level
}

Invoke-TestCase -Name 'The ring buffer keeps the most recent entries' -Body {
    $Script:RuntimeState.BufferMax = 5
    $Script:RuntimeState.Buffer.Clear()
    1..9 | ForEach-Object { Write-LogInfo "entry $_" -Component 'Unit' }
    $snap = @(Get-LogBufferSnapshot)
    Assert-Equal 5 $snap.Count
    Assert-Match 'entry 9' $snap[-1].Message
    $Script:RuntimeState.BufferMax = 1000
}

Invoke-TestCase -Name 'An unwritable log path does not abort the run' -Body {
    $Script:RuntimeState.LogPath = '/definitely/not/writable/x.log'
    $Script:RuntimeState.LogFailureWarned = $false
    Assert-NotThrows { Write-LogInfo 'this must not throw' -Component 'Unit' }
    $Script:RuntimeState.LogPath = $logFile
    $Script:RuntimeState.LogFailureWarned = $false
}

Invoke-TestCase -Name 'Level names map to increasing severity' -Body {
    $map = $Script:LogLevels
    Assert-True ($map['TRACE'] -lt $map['DEBUG'])
    Assert-True ($map['DEBUG'] -lt $map['INFO'])
    Assert-True ($map['INFO'] -lt $map['WARN'])
    Assert-True ($map['WARN'] -lt $map['ERROR'])
    Assert-True ($map['ERROR'] -lt $map['CRITICAL'])
}

Invoke-TestCase -Name 'Set-RuntimeLogLevel accepts every documented name and rejects none silently' -Body {
    foreach ($lvl in @('TRACE', 'DEBUG', 'INFO', 'WARN', 'ERROR', 'CRITICAL')) {
        Assert-NotNull (Set-RuntimeLogLevel -Level $lvl)
    }
    Assert-Equal 2 (Set-RuntimeLogLevel -Level 'BOGUS') 'an unknown level falls back to INFO'
    $Script:RuntimeState.MinLogLevelIndex = [int]$Script:LogLevels['INFO']
}

Start-TestSuite 'Log rotation and retention'

Invoke-TestCase -Name 'Rotation moves the active file to .1 when over the limit' -Body {
    $f = Join-Path $logDir 'rotate.log'
    Set-Content -LiteralPath $f -Value ('x' * 200000) -Encoding UTF8
    Assert-True (Invoke-LogRotation -Path $f -MaxSizeMB 0.1 -MaxFileCount 3)
    Assert-True (Test-Path -LiteralPath "$f.1")
    Assert-False (Test-Path -LiteralPath $f)
}

Invoke-TestCase -Name 'Rotation is a no-op under the limit' -Body {
    $f = Join-Path $logDir 'small.log'
    Set-Content -LiteralPath $f -Value 'tiny' -Encoding UTF8
    Assert-False (Invoke-LogRotation -Path $f -MaxSizeMB 50 -MaxFileCount 3)
    Assert-True (Test-Path -LiteralPath $f)
}

Invoke-TestCase -Name 'Rotation cascades generations and drops the oldest' -Body {
    $f = Join-Path $logDir 'cascade.log'
    foreach ($gen in 1..3) { Set-Content -LiteralPath "$f.$gen" -Value "gen$gen" -Encoding UTF8 }
    Set-Content -LiteralPath $f -Value ('y' * 200000) -Encoding UTF8
    $null = Invoke-LogRotation -Path $f -MaxSizeMB 0.1 -MaxFileCount 3
    Assert-False (Test-Path -LiteralPath "$f.4") 'generations beyond the max must be dropped'
    Assert-True (Test-Path -LiteralPath "$f.3")
}

Invoke-TestCase -Name 'Retention prunes only matching, only old, never the active log' -Body {
    $f = Join-Path $logDir 'retain.log'
    Set-Content -LiteralPath $f -Value 'active' -Encoding UTF8
    foreach ($n in @('retain.log.1', 'PrintSpoolerCleanup_20200101-000000.log')) {
        $old = Join-Path $logDir $n
        Set-Content -LiteralPath $old -Value 'old' -Encoding UTF8
        (Get-Item -LiteralPath $old).LastWriteTime = (Get-Date).AddDays(-400)
    }
    $keep = Join-Path $logDir 'UnrelatedNotes.txt'
    Set-Content -LiteralPath $keep -Value 'keep' -Encoding UTF8
    (Get-Item -LiteralPath $keep).LastWriteTime = (Get-Date).AddDays(-400)

    $pruned = Invoke-LogRetention -Path $f -RetentionDays 30
    Assert-GreaterThan $pruned 1
    Assert-True (Test-Path -LiteralPath $f)
    Assert-True (Test-Path -LiteralPath $keep)
}

Invoke-TestCase -Name 'Retention is disabled at 0 days' -Body {
    $f = Join-Path $logDir 'retain2.log'
    Set-Content -LiteralPath $f -Value 'x' -Encoding UTF8
    Assert-Equal 0 (Invoke-LogRetention -Path $f -RetentionDays 0)
}

Invoke-TestCase -Name 'A missing log path is tolerated' -Body {
    Assert-Equal 0 (Invoke-LogRetention -Path (Join-Path $logDir 'nope.log') -RetentionDays 30)
}

$Script:RuntimeState.LogPath = $savedLogPath
$Script:RuntimeState.JsonLog = $savedJson
Remove-Item -LiteralPath $logDir -Recurse -Force -ErrorAction SilentlyContinue

Start-TestSuite 'Log target resolution'

Invoke-TestCase -Name 'An explicit path is used and its directory is created' -Body {
    $dir = Join-Path ([IO.Path]::GetTempPath()) ("psq-mk-{0}" -f [Guid]::NewGuid().ToString('N'))
    $target = Join-Path $dir 'deep\nested.log'
    try {
        $r = Resolve-LogTarget -ExplicitPath $target
        Assert-True $r.Writable $r.Error
        Assert-Equal 'Parameter' $r.Source
        Assert-True (Test-Path -LiteralPath $dir)
    }
    finally { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue }
}

Invoke-TestCase -Name 'An unwritable explicit path reports the reason and stays writable=false' -Body {
    # The location must be one that genuinely CANNOT be written on the host running
    # the test. /proc is a Linux kernel filesystem, so on Windows the same string is
    # an ordinary relative path that opens fine - the assertion then failed for a
    # reason unrelated to the code under test. On Windows the equivalent structural
    # impossibility is naming a FILE as if it were a directory: no privilege level
    # makes that path writable.
    $unwritable = if ($IsWindows) {
        $parent = Join-Path ([IO.Path]::GetTempPath()) ("psq-not-a-dir-{0}" -f [Guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType File -Path $parent -Force
        Join-Path $parent 'cannot/write.log'
    }
    else {
        '/proc/1/cannot/write.log'
    }
    $r = Resolve-LogTarget -ExplicitPath $unwritable
    if ($r.Writable) { Assert-True $false "this location should not be writable: $unwritable" }
    else { Assert-NotNull $r.Error }
}

Invoke-TestCase -Name 'The probe leaves no blank record in the log' -Body {
    $f = Join-Path ([IO.Path]::GetTempPath()) ("psq-probe-{0}.log" -f [Guid]::NewGuid().ToString('N'))
    try {
        $r = Resolve-LogTarget -ExplicitPath $f
        Assert-True $r.Writable
        $content = @(Get-Content -LiteralPath $f -ErrorAction SilentlyContinue)
        Assert-Equal 0 @($content | Where-Object { $_.Trim() }).Count 'the writability probe must be silent'
    }
    finally { Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue }
}

Invoke-TestCase -Name 'New-ChildArgumentList rebuilds a child argument vector' -Body {
    # Regression for the `$args` collision: this helper used to build its result in
    # a variable named `$args`, which is the PowerShell AUTOMATIC variable holding
    # the current invocation's unbound arguments. It happened to work because the
    # function is always called with named parameters, but any future call passing
    # a positional argument would have had it silently swallowed - and the vector
    # would have come back containing the CALLER's arguments instead.
    $bound = [ordered]@{
        Force        = [switch]::new($true)
        Backup       = [switch]::new($true)
        OlderThan    = '2h'
        ThrottleLimit = 4
        Credential   = [pscredential]::new('u', (ConvertTo-SecureString 'p' -AsPlainText -Force))
        SelfElevate  = [switch]::new($true)
        PrinterName  = @('HP*', 'Canon*')
    }
    $argv = @(New-ChildArgumentList -BoundParameters $bound)

    Assert-True ($argv -contains '-Force') 'a present switch must become a bare -Name'
    Assert-True ($argv -contains '-Backup') 'a present switch must become a bare -Name'
    Assert-True ($argv -contains '-OlderThan') 'a scalar must emit its name'
    Assert-True ($argv -contains '2h') 'a scalar must emit its value'
    Assert-True ($argv -contains '-ThrottleLimit')
    Assert-True ($argv -contains '4')

    # Credentials must never cross a process boundary.
    Assert-False ($argv -contains '-Credential') 'PSCredential must NOT be forwarded to a child process'
    Assert-False ($argv -contains 'p') 'the credential secret must not appear in the argument vector'

    # -SelfElevate is excluded so the child does not try to elevate again.
    Assert-False ($argv -contains '-SelfElevate') '-SelfElevate must not be forwarded (infinite elevation loop)'

    # Arrays become a quoted comma list, and a value containing a space is quoted.
    Assert-True (($argv -join ' ') -match 'HP\*,Canon\*') ('arrays must join with commas: ' + ($argv -join ' '))
}

Invoke-TestCase -Name 'New-ChildArgumentList handles absent switches and an empty map' -Body {
    $off = [ordered]@{ Force = [switch]::new($false); Quiet = [switch]::new($false) }
    $argv = @(New-ChildArgumentList -BoundParameters $off)
    Assert-Equal 0 $argv.Count ('an absent switch must emit nothing: ' + ($argv -join ' '))

    $argv2 = @(New-ChildArgumentList -BoundParameters $null)
    Assert-Equal 0 $argv2.Count 'a null bound-parameter map must yield an empty vector'
}

Invoke-TestCase -Name 'A writable fallback root is always found' -Body {
    # Hermetic: Resolve-LogTarget CREATES each candidate directory as it probes
    # it, and one candidate is <script dir>\Logs - which, when the unit suite runs
    # against dist/, means Logs/ appears inside the release artifacts. Point
    # ScriptDirectory at a scratch directory for the duration of this test and
    # restore it afterwards, so the artifact tree stays clean and stage 8 can
    # assert that.
    $savedDir = $Script:RuntimeState.ScriptDirectory
    $scratch = Join-Path ([IO.Path]::GetTempPath()) ("psq-fallback-{0}" -f [Guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $scratch -Force
    $Script:RuntimeState.ScriptDirectory = $scratch
    try {
        $r = Resolve-LogTarget -ExplicitPath $null
        Assert-True $r.Writable $r.Error
        Assert-NotNull $r.Path
        Assert-True $r.Source 'a fallback root must report which candidate it used'
    }
    finally {
        $Script:RuntimeState.ScriptDirectory = $savedDir
        Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Start-TestSuite 'Console capability detection'

Invoke-TestCase -Name 'Capability detection never throws and always returns a shape' -Body {
    $c = Get-ConsoleCapability
    Assert-NotNull $c
    Assert-True ($c.PSObject.Properties.Name -contains 'SupportsColor')
    Assert-True ($c.PSObject.Properties.Name -contains 'SupportsUnicode')
    Assert-GreaterThan $c.BufferWidth 50
}

Invoke-TestCase -Name 'The ASCII palette covers every key of the Unicode palette' -Body {
    foreach ($k in $Script:GlyphUnicode.Keys) {
        Assert-True $Script:GlyphAscii.Contains($k) "ASCII palette is missing '$k'"
    }
}

Invoke-TestCase -Name 'Initialise-ConsolePalette honours -Ascii and -NoColor' -Body {
    $savedGlyphs = $Script:RuntimeState.Glyphs
    $null = Initialize-ConsolePalette -Ascii -NoColor
    Assert-True $Script:RuntimeState.UseAscii
    Assert-True $Script:RuntimeState.NoColor
    Assert-Equal $Script:GlyphAscii['Success'] $Script:RuntimeState.Glyphs['Success']
    $Script:RuntimeState.Glyphs = $savedGlyphs
    $Script:RuntimeState.UseAscii = $false
    $Script:RuntimeState.NoColor = $false
}

Start-TestSuite 'Selection description rendering'

Invoke-TestCase -Name 'No filters advertises the destructive default' -Body {
    Assert-Match 'ALL spool files' (Format-SelectionDescription -Arguments @{})
}

Invoke-TestCase -Name 'Each filter is described' -Body {
    Assert-Match "printers like 'HP\*'" (Format-SelectionDescription -Arguments @{ PrinterName = @('HP*') })
    Assert-Match 'excluding' (Format-SelectionDescription -Arguments @{ ExcludePrinterName = @('X*') })
    Assert-Match 'job id\(s\) 7, 8' (Format-SelectionDescription -Arguments @{ JobId = @(7, 8) })
    Assert-Match 'older than 2h' (Format-SelectionDescription -Arguments @{ OlderThan = '2h' })
    Assert-Match 'newer than 30m' (Format-SelectionDescription -Arguments @{ NewerThan = '30m' })
    Assert-Match 'larger than 1\.00 MB' (Format-SelectionDescription -Arguments @{ LargerThanBytes = 1MB })
    Assert-Match 'owner like' (Format-SelectionDescription -Arguments @{ OwnerName = 'CORP\*' })
}

Invoke-TestCase -Name 'Combined filters are all listed' -Body {
    $d = Format-SelectionDescription -Arguments @{ PrinterName = @('HP*'); OlderThan = '90m'; JobId = @(3) }
    Assert-Match 'jobs matching' $d
    Assert-Match 'HP\*' $d
    Assert-Match '90m' $d
}

# =============================================================================
#  Report schema contract
# -----------------------------------------------------------------------------
#  docs/Report-Schema.md is NORMATIVE: RMM tooling, jq and Python code against it.
#  Nothing else in the suite compared that document to the emitted bytes, so the
#  two drifted: the JSON emitted PascalCase keys inside a camelCase envelope (so
#  every documented field name resolved to nothing for a case-sensitive consumer),
#  and the payload dropped two documented job fields.
#
#  PowerShell cannot see this class of defect by property access, because
#  $obj.summary.Tool and $obj.summary.tool both work. Only the raw wire text
#  reveals it - so every assertion below matches the JSON TEXT case-sensitively.
# =============================================================================
Start-TestSuite 'Report schema contract (docs vs emitted JSON)'

$buildContractSummary = {
    # A graph populated with EVERY documented field, so "documented but absent"
    # cannot be explained away by an unpopulated fixture. Arrays are assigned, not
    # .Add()-ed: New-ResultObject initialises them as fixed-size @().
    $r = New-ResultObject -ComputerName 'SRV01' -Status 'Success'
    $r.IsLocal = $true
    $r.SpoolDirectory = 'C:\Windows\System32\spool\PRINTERS'
    $r.SpoolDirectorySource = 'Registry'
    $r.ServiceWasRunning = $true
    $r.ServiceIsRunning = $true
    $r.ServiceStopMs = 900
    $r.ServiceStartMs = 700
    $r.DependentServices = @('PrintNotify')
    $r.JobsFound = 1; $r.JobsSelected = 1; $r.JobsRemoved = 1
    $r.FilesFound = 2; $r.FilesSelected = 2; $r.FilesDeleted = 2
    $r.FilesFailed = 0; $r.FilesLocked = 0; $r.FilesAfter = 0
    $r.BytesFound = 4096; $r.BytesDeleted = 4096
    $r.ResurrectedFiles = 0
    $r.BackupPath = 'C:\Temp\bk.zip'; $r.BackupHash = ('a' * 64)
    $r.DryRun = $false; $r.RepairPerformed = $false; $r.DiagnosticsCollected = $true
    $r.Mode = 'Cleanup'
    $r.Warnings.Add('probe timed out') | Out-Null
    $r.Errors.Add('WinRM unreachable') | Out-Null

    $r.Jobs = @([pscustomobject]@{
            JobId = 7; PrinterName = 'HP LaserJet'; DocumentName = 'doc.pdf'; UserName = 'CORP\ann'
            SizeBytes = 2048; TotalPages = 4; PagesPrinted = 2; Status = 'Spooled'; Priority = 1
            SubmittedUtc = '2026-01-01T00:00:00Z'; Source = 'WMI+SHD'; ShdFile = '00007.SHD'; SplFile = '00007.SPL'
        })
    $r.RemovedFiles = @([pscustomobject]@{ file = '00007.SPL'; bytes = 2048; attempts = 1 })
    $r.FailedFiles = @([pscustomobject]@{ file = '00009.SPL'; bytes = 10; locked = $true; attempts = 3; error = 'in use' })
    $r.OrphanedFiles = @('00011.SHD')
    $r.ServiceTrace = @([pscustomobject]@{ service = 'Spooler'; action = 'Stop'; from = 'Running'; to = 'Stopped'; success = $true; timeUtc = '2026-01-01T00:00:01Z'; note = '' })
    $r.Diagnostics = [ordered]@{
        os = 'Windows Server 2022'; services = @(); spoolerProcess = $null; printers = @()
        drivers = @(); ports = @(); registry = @(); eventLog = @(); printEventLog = @()
        spoolDirectory = [ordered]@{
            path = 'C:\Windows\System32\spool\PRINTERS'; fileCount = 2; totalBytes = 4096
            byExtension = @([ordered]@{ ext = '.SPL'; count = 1; bytes = 2048 })
            oldestFile = [ordered]@{ name = '00007.SPL'; ageHours = 3.5 }
            newestFile = [ordered]@{ name = '00009.SPL'; ageHours = 0.5 }
        }
    }
    $r | Add-Member -NotePropertyName SelfTestPrint -NotePropertyValue ([pscustomobject]@{ attempted = $true; success = $true }) -Force

    New-RunSummary -Results @($r) -Arguments @{ Mode = 'Cleanup'; DryRun = $false } `
        -LogPath 'C:\logs\x.log' -ConfigSource 'Defaults' -Duration ([TimeSpan]::FromSeconds(9))
}

Invoke-TestCase -Name 'Every documented JSON field is emitted with that exact casing' -Body {
    $summary = & $buildContractSummary
    $json = ConvertTo-JsonReport -Summary $summary

    # Collect the keys the wire text actually carries, case-sensitively.
    $keys = @([regex]::Matches($json, '"([A-Za-z][A-Za-z0-9]*)"\s*:') |
        ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)

    # Collect documented field names, but ONLY from "| Field | Type |" tables.
    # The "| Value | Meaning |" table lists enum VALUES (WMI, SHD), not fields,
    # and including it produces false failures.
    $documented = New-Object System.Collections.Generic.List[string]
    $inFieldTable = $false
    foreach ($line in @(Get-Content -LiteralPath (Join-Path $Root 'docs/Report-Schema.md'))) {
        if ($line -notmatch '^\s*\|') { $inFieldTable = $false; continue }
        if ($line -match '^\s*\|[\s:\-\|]+\|?\s*$') { continue }        # delimiter row
        if ($line -match '^\|\s*Field\s*\|') { $inFieldTable = $true; continue }
        if ($line -match '^\|\s*Value\s*\|') { $inFieldTable = $false; continue }
        if (-not $inFieldTable) { continue }
        $firstCell = ($line -split '\|')[1]
        # `startedUtc` / `finishedUtc` share a row; `warnings[]` carries a suffix.
        foreach ($m in [regex]::Matches($firstCell, '`([A-Za-z][A-Za-z0-9]*)(?:\[\])?`')) {
            $documented.Add($m.Groups[1].Value)
        }
    }
    $documented = @($documented | Select-Object -Unique)
    Assert-GreaterThan $documented.Count 40 'the schema document should define a substantial contract'

    $missing = @($documented | Where-Object { $keys -cnotcontains $_ })
    Assert-Equal 0 $missing.Count $(if ($missing) {
            'documented but NOT emitted with that casing: ' + ($missing -join ', ')
        } else { "$($documented.Count) documented fields all present" })
}

Invoke-TestCase -Name 'No PascalCase key survives into the JSON report' -Body {
    # The envelope was always camelCase (schemaVersion, summary); the payload was
    # not, which is what made the document and the bytes disagree. One convention,
    # everywhere, is the only state that cannot drift.
    $summary = & $buildContractSummary
    $json = ConvertTo-JsonReport -Summary $summary
    $keys = @([regex]::Matches($json, '"([A-Za-z][A-Za-z0-9]*)"\s*:') |
        ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
    $pascal = @($keys | Where-Object { $_ -cmatch '^[A-Z]' })
    Assert-Equal 0 $pascal.Count $(if ($pascal) { 'PascalCase keys on the wire: ' + ($pascal -join ', ') } else { "$($keys.Count) keys, all camelCase" })
}

Invoke-TestCase -Name 'Every emitted key is documented' -Body {
    # The other direction: an undocumented field is an unpromised one, and
    # integrators will still come to depend on it.
    #
    # The exemption list is deliberately tiny and each entry is documented
    # SOMEWHERE in the schema file - just in prose or in the envelope code block
    # rather than in a "| Field |" table row.
    $proseDocumented = @('schemaVersion', 'summary', 'ext', 'count', 'bytes', 'name', 'ageHours',
        'path', 'fileCount', 'totalBytes', 'byExtension', 'oldestFile', 'newestFile',
        'file', 'locked', 'attempts', 'error', 'service', 'action', 'from', 'to', 'success', 'timeUtc', 'note',
        'attempted')
    $summary = & $buildContractSummary
    $json = ConvertTo-JsonReport -Summary $summary
    $keys = @([regex]::Matches($json, '"([A-Za-z][A-Za-z0-9]*)"\s*:') |
        ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)

    $doc = Get-Content -LiteralPath (Join-Path $Root 'docs/Report-Schema.md') -Raw
    # Array-valued fields are documented with their suffix - `jobs[]`,
    # `removedFiles[]`, `serviceTrace[]` - so a bare `jobs` between backticks
    # never matches. Accept either spelling.
    $undocumented = @($keys | Where-Object {
            $k = [regex]::Escape($_)
            ($doc -cnotmatch ('`' + $k + '`')) -and ($doc -cnotmatch ('`' + $k + '\[\]`')) -and
            ($proseDocumented -cnotcontains $_)
        })
    Assert-Equal 0 $undocumented.Count $(if ($undocumented) {
            'emitted but not documented anywhere: ' + ($undocumented -join ', ')
        } else { 'all emitted keys are documented' })
}

Invoke-TestCase -Name 'JSON values are never renamed, only keys' -Body {
    # The camelCase projection must touch keys alone. If it ever reached values,
    # status strings like "SuccessWithWarnings" would be corrupted and every
    # consumer comparison would break.
    $summary = & $buildContractSummary
    $json = ConvertTo-JsonReport -Summary $summary
    Assert-Match '"status":\s*"Success"' $json
    Assert-Match '"overallStatus":\s*"Success' $json
    Assert-Match '"source":\s*"WMI\+SHD"' $json
    Assert-Match '"computerName":\s*"SRV01"' $json
    Assert-Match '"printerName":\s*"HP LaserJet"' $json
}

Invoke-TestCase -Name 'Arrays stay arrays on the wire, including single-element ones' -Body {
    # ConvertTo-Json unrolls a single-element collection to a scalar unless the
    # producer wraps it. A consumer doing `.results[]` then breaks on the common
    # one-target run - which is the most common case of all.
    $summary = & $buildContractSummary
    $json = ConvertTo-JsonReport -Summary $summary
    Assert-Match '"results":\s*\[' $json
    Assert-Match '"jobs":\s*\[' $json
    Assert-Match '"dependentServices":\s*\[' $json
    Assert-Match '"orphanedFiles":\s*\[' $json
    Assert-Match '"serviceTrace":\s*\[' $json
    $obj = $json | ConvertFrom-Json
    Assert-Equal 1 @($obj.summary.results).Count
    Assert-Equal 1 @($obj.summary.results[0].jobs).Count
}

Invoke-TestCase -Name 'The documented jq-style projection resolves' -Body {
    # Straight from the schema document's consumption examples. If the casing is
    # wrong this returns nothing, which is precisely how the defect hid.
    $r1 = New-ResultObject -ComputerName 'SRV01' -Status 'Success'
    $r2 = New-ResultObject -ComputerName 'SRV02' -Status 'Failed'
    $summary = New-RunSummary -Results @($r1, $r2) -Arguments @{ Mode = 'Cleanup'; DryRun = $false } -Duration ([TimeSpan]::FromSeconds(1))
    $obj = (ConvertTo-JsonReport -Summary $summary) | ConvertFrom-Json
    $failed = @($obj.summary.results | Where-Object { $_.status -cne 'Success' } | ForEach-Object { $_.computerName })
    Assert-Equal 1 $failed.Count
    Assert-Equal 'SRV02' $failed[0]
}

# =============================================================================
#  Task Scheduler XML
# -----------------------------------------------------------------------------
#  A malformed task definition fails only at runtime, on Windows, when an operator
#  runs Install-SpoolerCleanupTask: schtasks.exe /create /xml rejects the file and
#  the entire scheduling feature is dead on arrival. Nothing else in the suite can
#  see it, because the Linux leg has no Task Scheduler.
#
#  Two schema details are easy to get wrong in a test and were gotten wrong once
#  already while writing this suite:
#    * the schema version is an ATTRIBUTE on <Task version="1.4">, not a child
#      element;
#    * <Subscription> may hold EITHER an XPath expression OR a <QueryList>
#      document - both are accepted by Task Scheduler, so asserting one form
#      reports a false defect.
# =============================================================================
Start-TestSuite 'Task Scheduler XML generation'

Invoke-TestCase -Name 'Every schedule type produces well-formed Task Scheduler 1.4 XML' -Body {
    foreach ($sch in @('Daily', 'Weekly', 'Hourly', 'AtStartup', 'OnEvent', 'Once')) {
        $xmlText = New-SpoolerTaskXml -Command 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe' `
            -Arguments '-NoProfile -ExecutionPolicy Bypass -File "C:\Tools\Clear-PrintSpoolerQueue.ps1" -Force -NonInteractive' `
            -Schedule $sch -EventLogName 'System' -EventSource 'Service Control Manager' -EventId 7031

        # NOTE: not `Assert-NotThrows { $doc = ... }`. That helper runs `& $Body`
        # in a CHILD scope, so the assignment never reaches this scope and $doc
        # stays $null - the test then fails on a scoping artefact rather than on
        # the XML. Capture the error instead.
        $doc = New-Object System.Xml.XmlDocument
        $parseError = $null
        try { $doc.LoadXml($xmlText) } catch { $parseError = $_.Exception.Message }
        Assert-Null $parseError "$sch must produce well-formed XML"
        Assert-NotNull $doc.DocumentElement "$sch parsed to a document element"

        $task = $doc.DocumentElement
        Assert-Equal 'Task' $task.LocalName "$sch root element"
        Assert-Equal 'http://schemas.microsoft.com/windows/2004/02/mit/task' $task.NamespaceURI "$sch namespace"
        # Version is an ATTRIBUTE, not a child element.
        Assert-Equal '1.4' $task.GetAttribute('version') "$sch schema version attribute"

        foreach ($el in @('RegistrationInfo', 'Triggers', 'Principals', 'Settings', 'Actions')) {
            Assert-NotNull $doc.SelectSingleNode("//*[local-name()='$el']") "$sch must contain <$el>"
        }
        Assert-Equal 1 @($doc.SelectNodes("//*[local-name()='Triggers']/*")).Count "$sch must declare exactly one trigger"
    }
}

Invoke-TestCase -Name 'Each schedule type emits the matching trigger element' -Body {
    $expected = @{
        'Daily' = 'CalendarTrigger'; 'Weekly' = 'CalendarTrigger'; 'Hourly' = 'CalendarTrigger'
        'AtStartup' = 'BootTrigger'; 'OnEvent' = 'EventTrigger'; 'Once' = 'TimeTrigger'
    }
    foreach ($sch in @($expected.Keys)) {
        $xmlText = New-SpoolerTaskXml -Command 'powershell.exe' -Arguments '-NoProfile' -Schedule $sch `
            -EventLogName 'System' -EventSource 'SCM' -EventId 7031
        $doc = New-Object System.Xml.XmlDocument; $doc.LoadXml($xmlText)
        $trig = $doc.SelectSingleNode("//*[local-name()='Triggers']/*")
        Assert-Equal $expected[$sch] $trig.LocalName "$sch trigger element"
    }
    # CalendarTrigger must carry the matching recurrence, or the task fires daily
    # regardless of what the operator asked for.
    $weekly = New-SpoolerTaskXml -Command 'powershell.exe' -Arguments '-NoProfile' -Schedule 'Weekly'
    Assert-Match 'ScheduleByWeek' $weekly
    $daily = New-SpoolerTaskXml -Command 'powershell.exe' -Arguments '-NoProfile' -Schedule 'Daily'
    Assert-Match 'ScheduleByDay' $daily
    $hourly = New-SpoolerTaskXml -Command 'powershell.exe' -Arguments '-NoProfile' -Schedule 'Hourly'
    Assert-Match 'Repetition' $hourly
}

Invoke-TestCase -Name 'An OnEvent trigger carries a usable subscription' -Body {
    # Without a Subscription the task registers but NEVER fires - a silent failure
    # the operator discovers only when the queue stays wedged.
    $xmlText = New-SpoolerTaskXml -Command 'powershell.exe' -Arguments '-NoProfile' -Schedule 'OnEvent' `
        -EventLogName 'System' -EventSource 'Service Control Manager' -EventId 7031
    $doc = New-Object System.Xml.XmlDocument; $doc.LoadXml($xmlText)
    $sub = $doc.SelectSingleNode("//*[local-name()='Subscription']")
    Assert-NotNull $sub 'EventTrigger must have a Subscription'

    $q = $sub.InnerText.Trim()
    Assert-True ($q.Length -gt 10) 'the subscription must not be empty'
    # Either an XPath expression or a QueryList document is acceptable; assert the
    # invariants that hold for BOTH rather than picking one form.
    if ($q.StartsWith('<')) {
        $d2 = New-Object System.Xml.XmlDocument
        $qErr = $null
        try { $d2.LoadXml($q) } catch { $qErr = $_.Exception.Message }
        Assert-Null $qErr 'QueryList subscription must itself be well-formed XML'
        Assert-Match 'QueryList' $q
    }
    else {
        $open = ([regex]::Matches($q, '\[')).Count
        $close = ([regex]::Matches($q, '\]')).Count
        Assert-Equal $open $close 'XPath subscription brackets must balance'
    }
    Assert-Match '7031' $q 'the subscription must filter on the requested event id'
    Assert-Match 'Service Control Manager' $q 'the subscription must filter on the requested source'
}

Invoke-TestCase -Name 'Hostile values cannot break the task XML' -Body {
    # The command line, arguments and description all land inside XML text nodes.
    # An unescaped & or < makes schtasks reject the whole file.
    $cases = @(
        @{ N = 'ampersand and embedded quotes'; A = '-File "C:\A & B\tool.ps1" -Note "say \"hi\""' }
        @{ N = 'literal XML markup in arguments'; A = '-Note <script>alert(1)</script>' }
        @{ N = 'angle brackets in working dir'; W = 'C:\Temp\<x>&y' }
        @{ N = 'non-ASCII description'; D = 'Обслуживание очереди печати — éè' }
        @{ N = 'empty command'; C = '' }
    )
    foreach ($c in $cases) {
        # Read every optional key through ContainsKey. Under Set-StrictMode a bare
        # `$c.A` on a hashtable that lacks the key THROWS rather than yielding
        # $null, so a case that omits 'A' would fail for a reason that has nothing
        # to do with XML escaping. This is why the suite is now pinned to Latest.
        $cmd = if ($c.ContainsKey('C')) { $c['C'] } else { 'powershell.exe' }
        $wd = if ($c.ContainsKey('W')) { $c['W'] } else { 'C:\Windows\System32' }
        $desc = if ($c.ContainsKey('D')) { $c['D'] } else { 'Automated maintenance.' }
        $argLine = if ($c.ContainsKey('A')) { $c['A'] } else { '' }
        $xmlText = New-SpoolerTaskXml -Command $cmd -Arguments $argLine -WorkingDirectory $wd -Description $desc -Schedule 'Daily'
        $doc = New-Object System.Xml.XmlDocument
        $parseError = $null
        try { $doc.LoadXml($xmlText) } catch { $parseError = $_.Exception.Message }
        Assert-Null $parseError "$($c.N) must still produce well-formed XML"
        Assert-NotNull $doc.DocumentElement $c.N
    }
}

Invoke-TestCase -Name 'Operator-supplied values survive the round trip verbatim' -Body {
    $marker = 'MARKER-' + [Guid]::NewGuid().ToString('N').Substring(0, 8)
    $xmlText = New-SpoolerTaskXml -Command 'C:\ps\pwsh.exe' -Arguments ("-Tag $marker") `
        -Schedule 'Daily' -StartTime '04:30'
    $doc = New-Object System.Xml.XmlDocument; $doc.LoadXml($xmlText)

    $argsNode = $doc.SelectSingleNode("//*[local-name()='Arguments']")
    Assert-NotNull $argsNode 'an Arguments element must exist'
    Assert-Match $marker $argsNode.InnerText 'arguments must be preserved verbatim'

    $cmdNode = $doc.SelectSingleNode("//*[local-name()='Command']")
    Assert-Equal 'C:\ps\pwsh.exe' $cmdNode.InnerText

    Assert-Match 'T04:30:00' $xmlText 'the requested start time must reach the trigger boundary'

    $userId = $doc.SelectSingleNode("//*[local-name()='UserId']")
    Assert-Equal 'NT AUTHORITY\SYSTEM' $userId.InnerText 'must run as SYSTEM by default'

    Assert-Match 'PT30M' $xmlText 'an execution time limit must be set, or a hung run blocks the next one'
}

Invoke-TestCase -Name 'The task XML survives the UTF-16LE encoding schtasks requires' -Body {
    # schtasks /create /xml requires UTF-16LE; writing UTF-8 makes it reject a
    # perfectly good definition with a confusing error.
    $xmlText = New-SpoolerTaskXml -Command 'powershell.exe' -Arguments '-NoProfile' -Schedule 'Daily'
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("psq-taskxml-{0}.xml" -f [Guid]::NewGuid().ToString('N'))
    try {
        [IO.File]::WriteAllText($tmp, $xmlText, [Text.Encoding]::Unicode)
        $bytes = [IO.File]::ReadAllBytes($tmp)
        Assert-True ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) 'UTF-16LE must carry a BOM'
        $back = [IO.File]::ReadAllText($tmp, [Text.Encoding]::Unicode)
        $doc = New-Object System.Xml.XmlDocument
        $parseError = $null
        try { $doc.LoadXml($back) } catch { $parseError = $_.Exception.Message }
        Assert-Null $parseError 'the UTF-16LE file must re-parse'
        Assert-Equal 'Task' $doc.DocumentElement.LocalName
    }
    finally { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
}

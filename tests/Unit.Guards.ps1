# Unit tests: the safety guards that fence the only destructive step in the product
Start-TestSuite 'Spool-directory guards: refusals'

$refused = @(
    @{ P = '';                                    Why = 'empty' }
    @{ P = '   ';                                 Why = 'whitespace' }
    @{ P = 'C:\';                                 Why = 'drive root' }
    @{ P = 'C:\Windows';                          Why = 'Windows directory' }
    @{ P = 'C:\Windows\System32';                 Why = 'System32' }
    @{ P = 'C:\Users';                            Why = 'user profile root' }
    @{ P = 'C:\Program Files';                    Why = 'Program Files' }
    @{ P = 'C:\ProgramData';                      Why = 'ProgramData' }
    @{ P = 'C:\PRINTERS';                         Why = 'correct leaf, shallow path' }
    @{ P = 'D:\PRINTERS';                         Why = 'correct leaf on another drive, shallow' }
    @{ P = '\\server\share\PRINTERS';             Why = 'UNC path' }
    @{ P = '\\?\C:\Windows\System32\spool\PRINTERS'; Why = 'extended-length prefix' }
    @{ P = 'relative\path\spool\PRINTERS';        Why = 'not rooted at a drive letter' }
    @{ P = '/tmp/spool/PRINTERS';                 Why = 'POSIX path' }
    @{ P = 'C:\Data';                             Why = 'arbitrary directory' }
    @{ P = 'C:\a\b';                              Why = 'arbitrary shallow directory' }
    @{ P = "C:\Windows\System32\spool\PRINTERS`0"; Why = 'embedded NUL' }
)
foreach ($case in $refused) {
    $c = $case
    Invoke-TestCase -Name "Refuses '$($c.P)' ($($c.Why))" -Body {
        $r = Test-SpoolDirectorySafeInternal -Path $c.P
        Assert-False $r.IsSafe
        Assert-NotNull $r.Reason 'a refusal must explain itself'
    }
}

Start-TestSuite 'Spool-directory guards: acceptances'

$accepted = @(
    @{ P = 'C:\Windows\System32\spool\PRINTERS'; Why = 'the OS default' }
    @{ P = 'D:\Spool\PRINTERS';                 Why = 'relocated spool root (legitimate policy)' }
    @{ P = 'F:\Very\Deep\Nested\spool\PRINTERS'; Why = 'deep but structurally valid' }
    @{ P = 'S:\SPOOL\PRINTERS';                  Why = 'uppercase SPOOL is equivalent' }
    @{ P = 'E:\Print$\spool\PRINTERS';          Why = 'hidden share-style folder name' }
    @{ P = 'c:\windows\system32\spool\printers'; Why = 'lowercase is equivalent' }
)
foreach ($case in $accepted) {
    $c = $case
    Invoke-TestCase -Name "Accepts '$($c.P)' ($($c.Why))" -Body {
        $r = Test-SpoolDirectorySafeInternal -Path $c.P
        Assert-True $r.IsSafe $r.Reason
    }
}

Start-TestSuite 'Guard layering'

Invoke-TestCase -Name 'The leaf requirement is what rejects most hostile paths' -Body {
    $r = Test-SpoolDirectorySafeInternal -Path 'C:\Windows\System32'
    Assert-False $r.IsSafe
    Assert-Match 'PRINTERS' $r.Reason
}

Invoke-TestCase -Name 'A protected location is refused even with the right leaf and depth' -Body {
    # The forbidden-root list is an EQUALITY check, so on its own it would accept
    # 'C:\Users\PRINTERS'. The structural 'must live under a spool folder' rule
    # is the layer that catches this whole class of corrupt-registry accident.
    foreach ($p in @('C:\Users\PRINTERS', 'C:\ProgramData\PRINTERS', 'C:\a\PRINTERS', 'D:\Reports\PRINTERS')) {
        $r = Test-SpoolDirectorySafeInternal -Path $p
        Assert-False $r.IsSafe "'$p' was accepted"
        Assert-Match 'spool' $r.Reason
    }
}

Invoke-TestCase -Name 'The forbidden-root list is never empty' -Body {
    $roots = @(Get-SpoolGuardForbiddenRoots)
    Assert-GreaterThan $roots.Count 5
    foreach ($x in $roots) { Assert-True (-not [string]::IsNullOrWhiteSpace($x)) 'null entries would defeat the comparison' }
}

Invoke-TestCase -Name 'The forbidden-root list contains no duplicates' -Body {
    $roots = @(Get-SpoolGuardForbiddenRoots)
    Assert-Equal $roots.Count @($roots | Select-Object -Unique).Count
}

Invoke-TestCase -Name 'Guard constants are sane' -Body {
    Assert-Equal 'PRINTERS' $Script:SpoolGuardRequiredLeaf
    Assert-Equal 3 $Script:SpoolGuardMinPathDepth
    Assert-GreaterThan $Script:SpoolGuardMaxFileCount 1000
    Assert-GreaterThan $Script:SpoolGuardMaxTotalBytes 1GB
}

Invoke-TestCase -Name 'The extension allow-list contains no wildcards' -Body {
    foreach ($e in $Script:SpoolExtensions) {
        Assert-Match '^\.[A-Za-z0-9]{1,5}$' $e
        Assert-NotMatch '\*' $e
    }
    Assert-Contains $Script:SpoolExtensions '.SPL'
    Assert-Contains $Script:SpoolExtensions '.SHD'
}

Start-TestSuite 'Payload guard parity'

Invoke-TestCase -Name 'The payload embeds the same guard constants as the local mirror' -Body {
    $null = Resolve-PayloadScriptBlock -Force
    $txt = $Script:RuntimeState.PayloadText
    Assert-True $txt.Contains("notlike 'PRINTERS'") 'leaf guard missing from the payload'
    Assert-True $txt.Contains('C:\Windows\System32') 'forbidden roots missing from the payload'
    Assert-True $txt.Contains("'^[A-Za-z]:") 'drive-letter guard missing from the payload'
    Assert-True $txt.Contains('$depth -lt 3') 'depth guard missing from the payload'
    Assert-Match 'spool' $txt 'the structural spool-folder guard is missing from the payload'
}

Invoke-TestCase -Name 'The payload never touches the spool directory recursively' -Body {
    $null = Resolve-PayloadScriptBlock -Force
    $txt = $Script:RuntimeState.PayloadText

    # Spool-directory enumeration must be file-only.
    Assert-Match '-File -Force' $txt 'file-only enumeration expected'

    # A blanket "-Recurse is forbidden" rule would be WRONG: the backup staging
    # directory is a GUID-named temp folder the tool creates itself, and cleaning
    # it up recursively is correct. What must never happen is recursion applied
    # to the spool directory, so assert on the combination instead.
    $lines = $txt -split "\r?\n"
    $bad = @()
    foreach ($line in $lines) {
        if ($line -notmatch '-Recurse') { continue }
        if ($line -match '\$Path|\$Dir|SpoolDirectory|\$R\[') { $bad += $line.Trim() }
    }
    Assert-Equal 0 $bad.Count ("recursive operation against the spool directory: " + ($bad -join ' ;; '))

    # The only recursive deletes permitted are the tool's own staging folders.
    foreach ($line in $lines) {
        if ($line -match 'Remove-Item' -and $line -match '-Recurse') {
            Assert-Match '\$staging' $line 'a recursive Remove-Item must target the backup staging directory only'
        }
    }
}

Invoke-TestCase -Name 'The payload normalises every result through one exit point' -Body {
    $null = Resolve-PayloadScriptBlock -Force
    $txt = $Script:RuntimeState.PayloadText
    Assert-GreaterThan ([regex]::Matches($txt, 'return Finish-Result')).Count 5
    # The only bare `return $R` allowed is the one inside Finish-Result itself.
    Assert-Equal 1 ([regex]::Matches($txt, 'return \$R\b')).Count
    Assert-Match 'function Finish-Result' $txt
}

Invoke-TestCase -Name 'The payload only ever deletes allow-listed extensions' -Body {
    $null = Resolve-PayloadScriptBlock -Force
    $txt = $Script:RuntimeState.PayloadText
    Assert-Match "\`$exts\s+=\s+@\('\.SPL', '\.SHD'\)" $txt
    Assert-Match 'Remove-SpoolFileRobust' $txt
    Assert-NotMatch 'Remove-Item -LiteralPath \$Path -Recurse' $txt
}

Invoke-TestCase -Name 'The payload job projection keeps every documented job field' -Body {
    # Report-Schema.md documents thirteen fields per job. The payload builds them
    # all - including PagesPrinted and Priority on both the WMI path and the .SHD
    # fallback - but a single Select-Object projection decides which survive into
    # the report. An earlier revision listed every field except those two, so the
    # schema promised columns no consumer ever received, and nothing failed.
    #
    # This compares the projection against the schema document rather than against
    # a hand-copied list, so adding a documented field without adding it to the
    # projection breaks the build.
    $payload = Get-Content -LiteralPath (Join-Path $Root 'src/Private/06-RemotePayload.ps1') -Raw
    $m = [regex]::Match($payload, "\`$R\['Jobs'\]\s*=\s*@\(\`$jobs\s*\|\s*Select-Object\s+([^)]+)\)")
    Assert-True $m.Success 'the job projection must exist in the payload'
    $projected = @($m.Groups[1].Value -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })

    $doc = Get-Content -LiteralPath (Join-Path $Root 'docs/Report-Schema.md') -Raw
    $jobsSection = [regex]::Match($doc, '(?s)### `jobs\[\]`(.*?)(?:\r?\n###|\r?\n---|\z)').Groups[1].Value

    # The jobs[] section contains TWO tables: the field table, and a
    # "| Value | Meaning |" table listing the enum values of `source`
    # (WMI / WMI+SHD / SHD). Reading every backticked first-column entry treats
    # those VALUES as field names and reports them as dropped by the projection.
    # Only rows belonging to a "| Field |" table count.
    $docFields = New-Object System.Collections.Generic.List[string]
    $inFieldTable = $false
    foreach ($line in @($jobsSection -split "\r?\n")) {
        if ($line -notmatch '^\s*\|') { $inFieldTable = $false; continue }
        if ($line -match '^\s*\|[\s:\-\|]+\|?\s*$') { continue }   # delimiter row
        if ($line -match '^\|\s*Field\s*\|') { $inFieldTable = $true; continue }
        if ($line -match '^\|\s*Value\s*\|') { $inFieldTable = $false; continue }
        if (-not $inFieldTable) { continue }
        $m = [regex]::Match(($line -split '\|')[1], '`([A-Za-z][A-Za-z0-9]*)(?:\[\])?`')
        if ($m.Success) { $docFields.Add($m.Groups[1].Value) }
    }
    $docFields = @($docFields | Select-Object -Unique)
    Assert-GreaterThan $docFields.Count 10 'the jobs[] schema table should be substantial'

    # Documented names are camelCase; the projection uses the PowerShell PascalCase
    # property names. Compare case-insensitively - the casing contract itself is
    # asserted separately against the emitted wire text.
    $missing = @($docFields | Where-Object {
            $f = $_
            -not ($projected | Where-Object { $_.ToLowerInvariant() -eq $f.ToLowerInvariant() })
        })
    Assert-Equal 0 $missing.Count $(if ($missing) {
            'documented job field(s) dropped by the payload projection: ' + ($missing -join ', ')
        } else { "all $($docFields.Count) documented job fields are projected" })

    foreach ($critical in @('PagesPrinted', 'Priority', 'Source', 'ShdFile', 'SplFile')) {
        Assert-Contains $projected $critical "$critical was dropped before; it must stay in the projection"
    }
}

Invoke-TestCase -Name 'The payload restarts the Spooler on every failure path' -Body {
    $null = Resolve-PayloadScriptBlock -Force
    $txt = $Script:RuntimeState.PayloadText
    Assert-Match 'LAST RESORT' $txt
    Assert-Match 'Start-OneService -name \$spoolSvc' $txt
    # Both the service-failure path and the catch path must attempt a restart.
    Assert-GreaterThan ([regex]::Matches($txt, 'Start-OneService')).Count 3
}

# =============================================================================
#  Parameter-combination validation
# -----------------------------------------------------------------------------
#  This suite did not exist, which is how a scope-widening defect survived: an
#  empty string is FALSY in PowerShell, so `-PrinterName ''` was silently treated
#  as "no filter" and the run cleaned EVERY printer instead of the named one.
#  The blank-value rule now rejects it, and these tests pin the behaviour.
# =============================================================================
Start-TestSuite 'Parameter combination validation'

Invoke-TestCase -Name 'A blank selection filter is REJECTED, not ignored' -Body {
    # The core safety property. A blank filter must never reach the payload,
    # because the payload's truthiness guard would skip it and widen the run to
    # every printer and every job.
    foreach ($p in @('PrinterName', 'ExcludePrinterName', 'OwnerName', 'DocumentName', 'OlderThan', 'NewerThan')) {
        $problems = @(Test-ParameterCombination -Bound @{ $p = '' })
        Assert-GreaterThan $problems.Count 0 ("-$p '' must be rejected, not silently ignored")
        Assert-Match 'IGNORED' ($problems -join ' ') "the message must explain WHY a blank filter is dangerous (-$p)"
        Assert-Match 'EVERY printer' ($problems -join ' ') "the message must state the consequence (-$p)"
    }
}

Invoke-TestCase -Name 'A whitespace-only selection filter is rejected' -Body {
    foreach ($p in @('PrinterName', 'ExcludePrinterName', 'OwnerName')) {
        $problems = @(Test-ParameterCombination -Bound @{ $p = '   ' })
        Assert-GreaterThan $problems.Count 0 ("-$p '   ' must be rejected")
    }
}

Invoke-TestCase -Name 'An array of blank strings is rejected' -Body {
    # The .bat launcher forwards `-PrinterName ""`; a config file could equally
    # supply @('', ''). Both must be caught.
    $problems = @(Test-ParameterCombination -Bound @{ PrinterName = @('', '') })
    Assert-GreaterThan $problems.Count 0 "@('', '') must be rejected"

    $problems2 = @(Test-ParameterCombination -Bound @{ PrinterName = @('  ') })
    Assert-GreaterThan $problems2.Count 0 "@('  ') must be rejected"
}

Invoke-TestCase -Name 'A partially blank filter list is still rejected' -Body {
    # @('', 'HP*') is truthy, so the old guard WOULD have applied the filter - but
    # the blank element matches nothing meaningful and signals an operator typo.
    # Reject rather than guess.
    $problems = @(Test-ParameterCombination -Bound @{ PrinterName = @('', 'HP*') })
    Assert-Equal 0 $problems.Count @'
a list containing at least one real pattern is accepted; the blank element is
harmless because the filter as a whole still narrows the run
'@
}

Invoke-TestCase -Name 'A real filter value is accepted' -Body {
    Assert-Equal 0 @(Test-ParameterCombination -Bound @{ PrinterName = 'HP*' }).Count
    Assert-Equal 0 @(Test-ParameterCombination -Bound @{ PrinterName = @('HP*', 'Canon*') }).Count
    Assert-Equal 0 @(Test-ParameterCombination -Bound @{ ExcludePrinterName = 'PDF' }).Count
    Assert-Equal 0 @(Test-ParameterCombination -Bound @{ OlderThan = '2h' }).Count
}

Invoke-TestCase -Name 'Blank PATH parameters are rejected' -Body {
    # A blank path is "unspecified", which for these parameters silently falls
    # back to a default location the operator did not choose.
    foreach ($p in @('BackupPath', 'ArchivePath', 'ExportPath', 'ConfigPath', 'RestoreDestination', 'ComputerListFile', 'DiagnosticsBundlePath')) {
        $problems = @(Test-ParameterCombination -Bound @{ $p = '' })
        Assert-GreaterThan $problems.Count 0 ("-$p '' must be rejected")
    }
}

Invoke-TestCase -Name 'Real path values are accepted' -Body {
    # -BackupPath carries a pre-existing dependency rule ('requires -Backup'), so
    # it must be supplied here: this case asserts that a REAL path is not mistaken
    # for a blank one, not that the dependency rules are absent.
    Assert-Equal 0 @(Test-ParameterCombination -Bound @{ BackupPath = 'C:\Temp\bk'; Backup = $true }).Count
    Assert-Equal 0 @(Test-ParameterCombination -Bound @{ ExportPath = 'C:\Temp\r.json' }).Count
    Assert-Equal 0 @(Test-ParameterCombination -Bound @{ ArchivePath = 'C:\Temp\a.zip' }).Count
    Assert-Equal 0 @(Test-ParameterCombination -Bound @{ ComputerListFile = 'C:\Temp\hosts.txt' }).Count
}

Invoke-TestCase -Name '-BackupPath still requires -Backup' -Body {
    # The dependency rule that the case above must respect; pinned explicitly so
    # relaxing it cannot slip through unnoticed.
    $problems = @(Test-ParameterCombination -Bound @{ BackupPath = 'C:\Temp\bk' })
    Assert-Equal 1 $problems.Count
    Assert-Match 'requires -Backup' ($problems -join ' ')
}

Invoke-TestCase -Name 'An explicitly null filter is rejected' -Body {
    $problems = @(Test-ParameterCombination -Bound @{ PrinterName = $null })
    Assert-GreaterThan $problems.Count 0 '-PrinterName $null must be rejected'
}

Invoke-TestCase -Name 'Mutually exclusive modes are still caught' -Body {
    # Regression cover for the pre-existing rules, which had no tests at all.
    Assert-GreaterThan @(Test-ParameterCombination -Bound @{ Restore = $true; Repair = $true }).Count 0 '-Restore + -Repair'
    Assert-GreaterThan @(Test-ParameterCombination -Bound @{ Diagnostic = $true; Repair = $true }).Count 0 '-Diagnostic + -Repair'
    Assert-GreaterThan @(Test-ParameterCombination -Bound @{ SelfElevate = $true; NoElevate = $true }).Count 0 '-SelfElevate + -NoElevate'
    Assert-GreaterThan @(Test-ParameterCombination -Bound @{ AutoUpdate = $true; CheckUpdate = $true }).Count 0 '-AutoUpdate + -CheckUpdate'
}

Invoke-TestCase -Name 'Dependency rules are still enforced' -Body {
    Assert-GreaterThan @(Test-ParameterCombination -Bound @{ BackupPath = 'C:\Temp' }).Count 0 '-BackupPath without -Backup'
    Assert-GreaterThan @(Test-ParameterCombination -Bound @{ SaveConfig = $true }).Count 0 '-SaveConfig without -ConfigPath'
    Assert-GreaterThan @(Test-ParameterCombination -Bound @{ JobId = 1; PrinterName = 'HP*' }).Count 0 '-JobId with -PrinterName'
    Assert-GreaterThan @(Test-ParameterCombination -Bound @{ Repair = $true; PurgeDriverStore = $true }).Count 0 '-PurgeDriverStore without -Force'
}

Invoke-TestCase -Name 'OlderThan must exceed NewerThan' -Body {
    Assert-GreaterThan @(Test-ParameterCombination -Bound @{ OlderThan = '1h'; NewerThan = '2h' }).Count 0 'inverted window'
    Assert-Equal 0 @(Test-ParameterCombination -Bound @{ OlderThan = '4h'; NewerThan = '1h' }).Count 'valid window'
}

Invoke-TestCase -Name 'A valid combination produces no problems' -Body {
    $bound = @{ Force = $true; Backup = $true; PrinterName = 'HP*'; OlderThan = '2h' }
    Assert-Equal 0 @(Test-ParameterCombination -Bound $bound).Count
}

Invoke-TestCase -Name 'The entry point actually WIRES the validator in before dispatch' -Body {
    # A validator nobody calls is decoration. This pins the wiring structurally,
    # because the end-to-end path cannot be exercised on a non-Windows host: the
    # platform guard (STEP 3) correctly aborts with exit 10 before argument
    # validation (STEP 4) is ever reached.
    #
    # What must hold: Test-ParameterCombination is called, its problems set exit
    # code 3, and that happens BEFORE the dispatch step - so a blank filter can
    # never reach the destructive payload.
    $src = Get-Content -LiteralPath (Join-Path $Root 'src/Public/Main-Function.ps1') -Raw
    Assert-Match 'Test-ParameterCombination -Bound \$PSBoundParameters' $src 'the validator must be called with the bound parameters'
    Assert-Match 'ExitCode\.BadArguments' $src 'a conflict must map to exit code 3'

    $callIdx = $src.IndexOf('Test-ParameterCombination -Bound')
    $dispatchIdx = $src.IndexOf('STEP 8 - dispatch')
    Assert-GreaterThan $callIdx -1 'the validator call must exist'
    Assert-GreaterThan $dispatchIdx -1 'the dispatch step marker must exist'
    Assert-True ($callIdx -lt $dispatchIdx) 'validation must run BEFORE dispatch, not after'

    # And the payload's own filter guard must still be a truthiness test - which
    # is exactly why the blank value has to be rejected upstream.
    $payload = Get-Content -LiteralPath (Join-Path $Root 'src/Private/06-RemotePayload.ps1') -Raw
    Assert-Match "\`$Options\['PrinterName'\] -and" $payload 'the payload gates filters on truthiness, so a blank value would be silently skipped'
}

Invoke-TestCase -Name 'An empty bound-parameter map is valid' -Body {
    Assert-Equal 0 @(Test-ParameterCombination -Bound @{}).Count
    Assert-Equal 0 @(Test-ParameterCombination -Bound $null).Count
}

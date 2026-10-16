<#
    StrictMode conformance suite.

    A great deal of PowerShell code "works" only because the host happens to run in
    lax mode. `Set-StrictMode -Version Latest` surfaces the class of defects that
    otherwise stay latent until a customer hardens their environment:

      * property access on $null - e.g. `Measure-Object` on an EMPTY pipeline
        returns $null, so `.Sum` throws
      * reading a property that does not exist on this particular object
      * uninitialised variables

    Both shapes occur naturally here: an empty result set is the normal case for a
    clean spool directory, and run results are deliberately heterogeneous (a
    Cleanup result carries JobsRemoved; a Diagnostic result does not).

    Every case in this file runs under Set-StrictMode -Version Latest and is turned
    back off in a finally block, so no other suite inherits the setting.
#>

Start-TestSuite 'StrictMode conformance'

function Script:Invoke-InStrictMode {
    [CmdletBinding()]
    param([Parameter(Mandatory)][scriptblock]$Body)
    Set-StrictMode -Version Latest
    try { & $Body }
    finally { Set-StrictMode -Off }
}

Invoke-TestCase -Name 'New-RunSummary tolerates an empty result set under StrictMode' -Body {
    # The defect that shipped: Measure-Object returns $null for an empty pipeline,
    # so `.Sum` threw "The property 'Sum' cannot be found on this object" - but
    # only when the caller had enabled StrictMode.
    $s = Invoke-InStrictMode { New-RunSummary -Results @() -Arguments @{ Force = $true } -Duration ([TimeSpan]::FromSeconds(1)) }
    Assert-NotNull $s
    Assert-Equal 'Failed' $s.OverallStatus
    # The property is TargetsAttempted, not TotalTargets - asserting a name that
    # does not exist returns $null in lax mode and silently "passes" as empty.
    Assert-Equal 0 $s.TargetsAttempted
    Assert-Equal 0 $s.JobsFound
    Assert-Equal 0 $s.JobsRemoved
    Assert-Equal 0 $s.BytesDeleted
    Assert-Equal 0 $s.Warnings
    Assert-Equal 0 $s.Errors
}

Invoke-TestCase -Name 'New-RunSummary sums a heterogeneous result set under StrictMode' -Body {
    # Records with DIFFERENT property shapes. Reading a missing property directly
    # throws under StrictMode, so aggregation must go through PSObject.Properties.
    $r = @(
        [pscustomobject]@{ ComputerName = 'A'; Status = 'Success'; JobsRemoved = 3; BytesDeleted = 1000; Warnings = @(); Errors = @() }
        [pscustomobject]@{ ComputerName = 'B'; Status = 'SuccessWithWarnings'; JobsRemoved = 1; BytesDeleted = 500; Warnings = @('w'); Errors = @() }
        # A Diagnostic result: no job or byte counters at all.
        [pscustomobject]@{ ComputerName = 'C'; Status = 'Diagnostic'; Warnings = @(); Errors = @() }
    )
    $s = Invoke-InStrictMode { New-RunSummary -Results $r -Arguments @{} -Duration ([TimeSpan]::FromSeconds(2)) }
    Assert-Equal 3 $s.TargetsAttempted
    Assert-Equal 4 $s.JobsRemoved
    Assert-Equal 1500 $s.BytesDeleted
    Assert-Equal 1 $s.Warnings
    Assert-Equal 0 $s.Errors
    Assert-Equal 'SuccessWithWarnings' $s.OverallStatus
}

Invoke-TestCase -Name 'New-RunSummary sums large byte counts without overflow under StrictMode' -Body {
    $big = [long]::MaxValue / 4
    $r = @(
        [pscustomobject]@{ ComputerName = 'A'; Status = 'Success'; BytesDeleted = $big }
        [pscustomobject]@{ ComputerName = 'B'; Status = 'Success'; BytesDeleted = $big }
    )
    $s = Invoke-InStrictMode { New-RunSummary -Results $r -Arguments @{} -Duration ([TimeSpan]::Zero) }
    Assert-Equal ([long]($big * 2)) $s.BytesDeleted
}

Invoke-TestCase -Name 'Report serialisers tolerate an empty summary under StrictMode' -Body {
    $empty = Invoke-InStrictMode { New-RunSummary -Results @() -Arguments @{} -Duration ([TimeSpan]::Zero) }
    Assert-Equal '' (Invoke-InStrictMode { ConvertTo-CsvReport -Summary $empty })
    Assert-True ((Invoke-InStrictMode { ConvertTo-JsonReport -Summary $empty }).Length -gt 20)
    Assert-True ((Invoke-InStrictMode { ConvertTo-HtmlReport -Summary $empty }).Length -gt 200)
    Assert-True ((Invoke-InStrictMode { ConvertTo-TextReport -Summary $empty }).Length -gt 20)
}

Invoke-TestCase -Name 'Write-ConsoleTable tolerates empty and single-row input under StrictMode' -Body {
    # Column widths are summed from a hashtable of measured widths; an empty input
    # set leaves it empty, and Measure-Object would have returned $null.
    Invoke-InStrictMode { @() | Write-ConsoleTable -Property @('Name') }
    Invoke-InStrictMode { @(Get-Process -Id $PID) | Write-ConsoleTable -Property @('ProcessName') }
    Assert-True $true 'Write-ConsoleTable survived both shapes without throwing'
}

Invoke-TestCase -Name 'Diagnostic spool-directory sizing tolerates an empty directory under StrictMode' -Body {
    # The payload summed file sizes with Measure-Object. An EMPTY spool directory -
    # the normal, healthy case - produced a $null measure result and then threw.
    $empty = Join-Path ([IO.Path]::GetTempPath()) ("psq-empty-spool-{0}" -f [Guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $empty -Force
    try {
        $bytes = Invoke-InStrictMode {
            $files = @(Get-ChildItem -LiteralPath $empty -File -Force -ErrorAction SilentlyContinue)
            $t = [long]0
            foreach ($f in $files) { $t += [long]$f.Length }
            $t
        }
        Assert-Equal 0 $bytes
    }
    finally { Remove-Item -LiteralPath $empty -Recurse -Force -ErrorAction SilentlyContinue }
}

Invoke-TestCase -Name 'Health aggregation tolerates an unavailable spooler under StrictMode' -Body {
    $r = Invoke-InStrictMode { Test-PrintSpoolerHealth }
    Assert-True (@($r).Count -ge 1)
    Assert-NotNull @($r)[0].Health
}

Invoke-TestCase -Name 'Get-ResultPropertyValue is StrictMode-safe for absent properties' -Body {
    $o = [pscustomobject]@{ Present = 7 }
    Assert-Equal 7 (Invoke-InStrictMode { Get-ResultPropertyValue -Item $o -Name 'Present' })
    # Reading $o.Absent directly would throw under StrictMode; the helper must not.
    Assert-Null (Invoke-InStrictMode { Get-ResultPropertyValue -Item $o -Name 'Absent' })
    Assert-Null (Invoke-InStrictMode { Get-ResultPropertyValue -Item $null -Name 'Present' })
}

Invoke-TestCase -Name 'Version comparison tolerates null inputs under StrictMode' -Body {
    # Compare-ProductVersion returns -1 / 0 / +1 and its parameters are
    # -Reference / -Difference, both [AllowNull()][AllowEmptyString()].
    Assert-Equal 1 (Invoke-InStrictMode { Compare-ProductVersion -Reference '1.0.0' -Difference $null })
    Assert-Equal 0 (Invoke-InStrictMode { Compare-ProductVersion -Reference '1.2.3' -Difference '1.2.3' })
    Assert-Equal -1 (Invoke-InStrictMode { Compare-ProductVersion -Reference '1.0.0' -Difference '2.0.0' })
}

Invoke-TestCase -Name 'Configuration merge tolerates an absent config file under StrictMode' -Body {
    $c = Invoke-InStrictMode { Get-EffectiveConfiguration -ScriptDirectory (Join-Path ([IO.Path]::GetTempPath()) 'no-such-dir') -NoConfigFile }
    Assert-True ($c.Config.Count -ge 7)  # Service, Cleanup, Logging, Execution, Update, Diagnostics, Remote
    # Test-Configuration returns an ARRAY OF PROBLEM STRINGS; empty means valid.
    $problems = Invoke-InStrictMode { Test-Configuration -Configuration $c.Config }
    Assert-Equal 0 @($problems | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count (@($problems) -join '; ')
}

Invoke-TestCase -Name 'Formatters tolerate zero and null under StrictMode' -Body {
    Assert-Equal '0 B' (Invoke-InStrictMode { Format-ByteSize -Bytes 0 })
    Assert-Equal '0 B' (Invoke-InStrictMode { Format-ByteSize -Bytes $null })
    # Format-Duration renders "HH:mm:ss.fff" below a day and "Nd HHh MMm SSs" above.
    Assert-Equal '00:00:00.005' (Invoke-InStrictMode { Format-Duration -Duration ([TimeSpan]::FromMilliseconds(5)) })
    Assert-Equal '00:00:00.000' (Invoke-InStrictMode { Format-Duration -Duration $null })
    Assert-Match '^1d ' (Invoke-InStrictMode { Format-Duration -Duration ([TimeSpan]::FromHours(26)) })
}

Invoke-TestCase -Name 'Update-state read/write tolerates a missing file under StrictMode' -Body {
    $p = Join-Path ([IO.Path]::GetTempPath()) ("psq-sm-state-{0}.json" -f [Guid]::NewGuid().ToString('N'))
    try {
        $st = Invoke-InStrictMode { Read-UpdateState -Path $p }
        Assert-NotNull $st
        Assert-True ($st -is [System.Collections.IDictionary])
    }
    finally { Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue }
}

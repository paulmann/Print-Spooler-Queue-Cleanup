<#
.SYNOPSIS
    Minimal, dependency-free PowerShell test harness.

.DESCRIPTION
    The product's own rule is "native Windows utilities only - no external
    modules", and the test suite follows the same rule: it must run on a clean
    Windows machine with nothing installed, and on a Linux CI runner with only
    PowerShell present. Requiring Pester would break that promise for no real
    gain at this scale.

    Assertions throw on failure, which the runner converts into a reported FAIL
    with the offending expression and the call site.

.NOTES
    Dot-source this file, then use Assert-* inside Script:TestCases entries.
#>

$Script:TestPasses  = 0
$Script:TestFailures = 0
$Script:TestResults  = New-Object System.Collections.Generic.List[object]
$Script:CurrentSuite = ''

function Script:Start-TestSuite {
    param([Parameter(Mandatory)][string]$Name)
    $Script:CurrentSuite = $Name
    if (-not $env:PSQ_TEST_QUIET) { Write-Host "`n--- $Name ---" -ForegroundColor Cyan }
}

function Script:Add-TestResult {
    param([string]$Suite, [string]$Name, [bool]$Passed, [string]$Detail)
    $Script:TestResults.Add([pscustomobject]@{
        Suite = $Suite; Name = $Name; Result = $(if ($Passed) { 'PASS' } else { 'FAIL' }); Detail = $Detail
    }) | Out-Null
    if ($Passed) { $Script:TestPasses++ } else { $Script:TestFailures++ }
    if (-not $env:PSQ_TEST_QUIET) {
        if ($Passed) { Write-Host ("  PASS  {0}" -f $Name) -ForegroundColor Green }
        else { Write-Host ("  FAIL  {0}`n        {1}" -f $Name, $Detail) -ForegroundColor Red }
    }
}

function Script:Invoke-TestCase {
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][scriptblock]$Body)
    try {
        & $Body
        Add-TestResult -Suite $Script:CurrentSuite -Name $Name -Passed $true -Detail ''
    }
    catch {
        $detail = $_.Exception.Message
        if ($_.ScriptStackTrace) { $detail += "  [$($_.ScriptStackTrace -split "`n" | Select-Object -First 1)]" }
        Add-TestResult -Suite $Script:CurrentSuite -Name $Name -Passed $false -Detail $detail
    }
}

# ---------------------------------------------------------------------------
#  Assertions
# ---------------------------------------------------------------------------
function Script:Assert-True {
    param([Parameter(Position = 0, ValueFromPipeline = $true)]$Value, [Parameter(Position = 1)][string]$Because = '')
    if (-not $Value) { throw ("Expected a truthy value but got '{0}'. {1}" -f $Value, $Because) }
}

function Script:Assert-False {
    param([Parameter(Position = 0, ValueFromPipeline = $true)]$Value, [Parameter(Position = 1)][string]$Because = '')
    if ($Value) { throw ("Expected a falsy value but got '{0}'. {1}" -f $Value, $Because) }
}

function Script:Assert-Equal {
    param([Parameter(Position = 0)]$Expected, [Parameter(Position = 1)]$Actual, [Parameter(Position = 2)][AllowEmptyString()][string]$Because = '')
    if ($Expected -is [string] -or $Actual -is [string]) {
        if ([string]$Expected -cne [string]$Actual) { throw ("Expected '{0}' but got '{1}'. {2}" -f $Expected, $Actual, $Because) }
        return
    }
    if ($Expected -ne $Actual) { throw ("Expected '{0}' but got '{1}'. {2}" -f $Expected, $Actual, $Because) }
}

function Script:Assert-NotEqual {
    param([Parameter(Position = 0)]$NotExpected, [Parameter(Position = 1)]$Actual, [Parameter(Position = 2)][AllowEmptyString()][string]$Because = '')
    if ($NotExpected -eq $Actual) { throw ("Expected anything but '{0}'. {1}" -f $NotExpected, $Because) }
}

function Script:Assert-Null {
    param([Parameter(Position = 0, ValueFromPipeline = $true)]$Value, [Parameter(Position = 1)][string]$Because = '')
    if ($null -ne $Value) { throw ("Expected null but got '{0}'. {1}" -f $Value, $Because) }
}

function Script:Assert-NotNull {
    param([Parameter(Position = 0, ValueFromPipeline = $true)]$Value, [Parameter(Position = 1)][string]$Because = '')
    if ($null -eq $Value) { throw ("Expected a value but got null. {0}" -f $Because) }
}

function Script:Assert-Match {
    param([Parameter(Position = 0)][string]$Pattern, [Parameter(Position = 1)][AllowNull()][AllowEmptyString()][string]$Value, [Parameter(Position = 2)][AllowEmptyString()][string]$Because = '')
    if ($Value -notmatch $Pattern) { throw ("Expected a match for /{0}/ in '{1}'. {2}" -f $Pattern, $Value, $Because) }
}

function Script:Assert-NotMatch {
    param([Parameter(Position = 0)][string]$Pattern, [Parameter(Position = 1)][AllowNull()][AllowEmptyString()][string]$Value, [Parameter(Position = 2)][AllowEmptyString()][string]$Because = '')
    if ($Value -match $Pattern) { throw ("Expected NO match for /{0}/ but '{1}' matched. {2}" -f $Pattern, $Value, $Because) }
}

function Script:Assert-Contains {
    param([Parameter(Position = 0)]$Collection, [Parameter(Position = 1)]$Item, [Parameter(Position = 2)][AllowEmptyString()][string]$Because = '')
    if (@($Collection) -notcontains $Item) { throw ("Expected the collection to contain '{0}'. {1}" -f $Item, $Because) }
}

function Script:Assert-NotContains {
    param([Parameter(Position = 0)]$Collection, [Parameter(Position = 1)]$Item, [Parameter(Position = 2)][AllowEmptyString()][string]$Because = '')
    if (@($Collection) -contains $Item) { throw ("Expected the collection NOT to contain '{0}'. {1}" -f $Item, $Because) }
}

function Script:Assert-Count {
    param([Parameter(Position = 0)][int]$Expected, [Parameter(Position = 1)]$Collection, [Parameter(Position = 2)][AllowEmptyString()][string]$Because = '')
    $actual = @($Collection).Count
    if ($actual -ne $Expected) { throw ("Expected {0} item(s) but got {1}. {2}" -f $Expected, $actual, $Because) }
}

function Script:Assert-Throws {
    param([Parameter(Mandatory, Position = 0)][scriptblock]$Body, [Parameter(Position = 1)][AllowEmptyString()][string]$Because = '')
    $threw = $false
    try { & $Body | Out-Null } catch { $threw = $true }
    if (-not $threw) { throw ("Expected an exception but none was thrown. {0}" -f $Because) }
}

function Script:Assert-NotThrows {
    param([Parameter(Mandatory, Position = 0)][scriptblock]$Body, [Parameter(Position = 1)][AllowEmptyString()][string]$Because = '')
    try { & $Body | Out-Null }
    catch { throw ("Expected no exception but got: {0}. {1}" -f $_.Exception.Message, $Because) }
}

function Script:Assert-GreaterThan {
    param([Parameter(Position = 0)]$Value, [Parameter(Position = 1)]$Threshold, [Parameter(Position = 2)][AllowEmptyString()][string]$Because = '')
    if (-not ([double]$Value -gt [double]$Threshold)) { throw ("Expected {0} > {1}. {2}" -f $Value, $Threshold, $Because) }
}

function Script:Invoke-WithEnvironment {
    <#
    .SYNOPSIS
        Sets environment variables for the duration of a scriptblock, then restores
        the EXACT previous state.

    .DESCRIPTION
        Tests must never `Remove-Item Env:\X` for a variable they did not create.
        The quality gate sets PSQ_LOG_PATH to keep every stage from writing logs
        into the release artifact tree; a test that deleted it would silently
        un-redirect logging for every later stage, which presents as a baffling
        failure in an unrelated assertion. Save-and-restore is the only safe shape.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Variables,
        [Parameter(Mandatory)][scriptblock]$Body
    )

    $saved = @{}
    foreach ($key in $Variables.Keys) {
        $path = "Env:\$key"
        if (Test-Path -LiteralPath $path) {
            $saved[$key] = [pscustomobject]@{ Existed = $true; Value = (Get-Item -LiteralPath $path).Value }
        }
        else {
            $saved[$key] = [pscustomobject]@{ Existed = $false; Value = $null }
        }
        Set-Item -LiteralPath $path -Value $Variables[$key]
    }
    try {
        & $Body
    }
    finally {
        foreach ($key in $saved.Keys) {
            $path = "Env:\$key"
            if ($saved[$key].Existed) { Set-Item -LiteralPath $path -Value $saved[$key].Value }
            else { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue }
        }
    }
}

function Script:Get-TestSummary {
    return [pscustomobject]@{
        Passed   = $Script:TestPasses
        Failed   = $Script:TestFailures
        Total    = $Script:TestPasses + $Script:TestFailures
        Results  = $Script:TestResults.ToArray()
    }
}

function Script:Reset-TestCounters {
    $Script:TestPasses = 0
    $Script:TestFailures = 0
    $Script:TestResults = New-Object System.Collections.Generic.List[object]
}

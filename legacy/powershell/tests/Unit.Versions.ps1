# Unit tests: version parsing and comparison (the updater's decision logic)
Start-TestSuite 'Version parsing'

Invoke-TestCase -Name 'Parses a plain three-part version' -Body {
    $v = ConvertTo-SemVer -Value '4.0.0'
    Assert-NotNull $v
    Assert-Equal '4.0.0.0' $v.Version.ToString()
    Assert-False $v.IsPrerelease
}

Invoke-TestCase -Name 'Parses a v-prefixed tag' -Body {
    Assert-Equal '4.1.2.0' (ConvertTo-SemVer -Value 'v4.1.2').Version.ToString()
}

Invoke-TestCase -Name 'Parses a two-part version by zero-padding' -Body {
    Assert-Equal '4.0.0.0' (ConvertTo-SemVer -Value '4.0').Version.ToString()
}

Invoke-TestCase -Name 'Splits prerelease metadata off the numeric core' -Body {
    $v = ConvertTo-SemVer -Value '4.0.0-beta.1'
    Assert-True $v.IsPrerelease
    Assert-Equal 'beta.1' $v.Label
    Assert-Equal '4.0.0.0' $v.Version.ToString()
}

Invoke-TestCase -Name 'Ignores build metadata after +' -Body {
    $v = ConvertTo-SemVer -Value '4.0.0+build.99'
    Assert-Equal '4.0.0.0' $v.Version.ToString()
    Assert-False $v.IsPrerelease
}

Invoke-TestCase -Name 'Rejects non-numeric components' -Body {
    Assert-Null (ConvertTo-SemVer -Value 'latest')
    Assert-Null (ConvertTo-SemVer -Value '4.x.0')
    Assert-Null (ConvertTo-SemVer -Value '')
    Assert-Null (ConvertTo-SemVer -Value $null)
}

Invoke-TestCase -Name 'Truncates more than four components rather than throwing' -Body {
    Assert-NotNull (ConvertTo-SemVer -Value '1.2.3.4.5')
}

Start-TestSuite 'Version comparison'

$compareCases = @(
    @{ A = '3.1.0';      B = '4.0.0';      E = -1; Why = 'major upgrade' }
    @{ A = '4.0.0';      B = '4.0.0';      E = 0;  Why = 'identical' }
    @{ A = '4.0.1';      B = '4.0.0';      E = 1;  Why = 'patch ahead' }
    @{ A = '4.10.0';     B = '4.9.0';      E = 1;  Why = 'numeric, not lexical, comparison' }
    @{ A = '4.0.0';      B = 'v4.0.0';     E = 0;  Why = 'v prefix is normalised' }
    @{ A = '4.0';        B = '4.0.0.0';    E = 0;  Why = 'missing parts are zero' }
    @{ A = '4.0.0-beta'; B = '4.0.0';      E = -1; Why = 'prerelease sorts before release' }
    @{ A = '4.0.0';      B = '4.0.0-beta'; E = 1;  Why = 'release sorts after prerelease' }
    @{ A = '4.0.0-alpha';B = '4.0.0-beta'; E = -1; Why = 'prerelease labels compare ordinally' }
    @{ A = '';           B = '4.0.0';      E = -1; Why = 'unknown is older than known' }
    @{ A = '4.0.0';      B = '';           E = 1;  Why = 'known is newer than unknown' }
    @{ A = '';           B = '';           E = 0;  Why = 'both unknown are equal' }
    @{ A = 'garbage';    B = '4.0.0';      E = -1; Why = 'unparseable is treated as older' }
    @{ A = '4.0.0';      B = 'garbage';    E = 1;  Why = 'unparseable right side is older' }
    @{ A = '10.0.0';     B = '9.9.9';      E = 1;  Why = 'double-digit major' }
)
foreach ($c in $compareCases) {
    $case = $c
    Invoke-TestCase -Name "Compare '$($case.A)' vs '$($case.B)' ($($case.Why))" -Body {
        Assert-Equal $case.E (Compare-ProductVersion -Reference $case.A -Difference $case.B)
    }
}

Invoke-TestCase -Name 'Comparison is antisymmetric' -Body {
    foreach ($pair in @(@('1.0.0','2.0.0'), @('4.0.0-beta','4.0.0'), @('4.10.0','4.9.0'))) {
        $fwd = Compare-ProductVersion -Reference $pair[0] -Difference $pair[1]
        $rev = Compare-ProductVersion -Reference $pair[1] -Difference $pair[0]
        Assert-Equal $fwd (-$rev) "asymmetry for $($pair -join ' / ')"
    }
}

Invoke-TestCase -Name 'Comparison is reflexive for every shape' -Body {
    foreach ($v in @('1.0.0', '4.0.0-beta', 'v2.3', '4.0.0+build')) {
        Assert-Equal 0 (Compare-ProductVersion -Reference $v -Difference $v) $v
    }
}

Start-TestSuite 'Update decision (Test-NewerVersion)'

Invoke-TestCase -Name 'A higher stable version is an update' -Body {
    Assert-True (Test-NewerVersion -Current '4.0.0' -Candidate '4.0.1')
}

Invoke-TestCase -Name 'The same version is not an update' -Body {
    Assert-False (Test-NewerVersion -Current '4.0.0' -Candidate '4.0.0')
}

Invoke-TestCase -Name 'An older version is not an update' -Body {
    Assert-False (Test-NewerVersion -Current '4.0.0' -Candidate '3.9.9')
}

Invoke-TestCase -Name 'Prereleases are ignored by default' -Body {
    Assert-False (Test-NewerVersion -Current '4.0.0' -Candidate '4.1.0-rc1')
}

Invoke-TestCase -Name 'Prereleases are accepted when explicitly requested' -Body {
    Assert-True (Test-NewerVersion -Current '4.0.0' -Candidate '4.1.0-rc1' -IncludePrerelease)
}

Invoke-TestCase -Name 'An unparseable candidate is never an update' -Body {
    Assert-False (Test-NewerVersion -Current '4.0.0' -Candidate 'not-a-version')
    Assert-False (Test-NewerVersion -Current '4.0.0' -Candidate '')
    Assert-False (Test-NewerVersion -Current '4.0.0' -Candidate $null)
}

Start-TestSuite 'Version stamp extraction from a remote script'

Invoke-TestCase -Name 'Reads the stamped version from source text' -Body {
    $sample = @'
$Script:ScriptVersion = '9.8.7'
function Foo { }
'@
    Assert-Equal '9.8.7' (Get-VersionFromRemoteScript -Content $sample)
}

Invoke-TestCase -Name 'Prefers the AST assignment over a comment mention' -Body {
    $sample = @'
<#
    Script Version: 1.1.1 (old note in a comment)
#>
$Script:ScriptVersion = '5.4.3'
'@
    Assert-Equal '5.4.3' (Get-VersionFromRemoteScript -Content $sample)
}

Invoke-TestCase -Name 'Falls back to the NOTES-style stamp when no assignment exists' -Body {
    $sample = "<#`n.NOTES`n    Script Version: 2.2.2`n#>`n"
    Assert-Equal '2.2.2' (Get-VersionFromRemoteScript -Content $sample)
}

Invoke-TestCase -Name 'Returns null for content with no version at all' -Body {
    Assert-Null (Get-VersionFromRemoteScript -Content 'Write-Host "hello"')
    Assert-Null (Get-VersionFromRemoteScript -Content '')
    Assert-Null (Get-VersionFromRemoteScript -Content $null)
}

Invoke-TestCase -Name 'The live built artifact reports its own version' -Body {
    $text = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'dist/Clear-PrintSpoolerQueue.ps1') -Raw
    Assert-Equal $Script:ScriptVersion (Get-VersionFromRemoteScript -Content $text)
}

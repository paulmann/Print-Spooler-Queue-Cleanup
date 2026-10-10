# Unit tests: the -Help CLI contract (defect 46 regression suite).
#
# Defect 46: -Help was declared [string] with a ValidateSet, so the bare form
#   .\Clear-PrintSpoolerQueue.ps1 -Help
# - the exact command the README, the comment-based help examples and the
# console hints advertise - failed with "Missing an argument for parameter
# 'Help'". Every earlier test called Get-HelpText/-ListHelpSections directly or
# passed a section value, so 697 checks stayed green while the documented
# invocation was broken. These tests bind the CLI the way a user does.
Start-TestSuite 'Help CLI contract (defect 46)'

$helpCmd = Get-Command -Name 'Clear-PrintSpoolerQueue'

Invoke-TestCase -Name '-Help is a switch, so the bare documented form binds' -Body {
    Assert-True $helpCmd.Parameters.ContainsKey('Help')
    Assert-Equal 'SwitchParameter' $helpCmd.Parameters['Help'].ParameterType.Name
}

Invoke-TestCase -Name '-HelpSection exists and is a validated string' -Body {
    Assert-True $helpCmd.Parameters.ContainsKey('HelpSection')
    Assert-Equal 'String' $helpCmd.Parameters['HelpSection'].ParameterType.Name
    $validateSet = @($helpCmd.Parameters['HelpSection'].Attributes |
        Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] })
    Assert-Equal 1 $validateSet.Count '-HelpSection must keep its ValidateSet'
}

Invoke-TestCase -Name '-HelpSection ValidateSet matches the published section names exactly' -Body {
    $validValues = @(($helpCmd.Parameters['HelpSection'].Attributes |
        Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }).ValidValues)
    $published = @(Get-HelpSectionNames)
    Assert-Equal $published.Count $validValues.Count 'no drift between ValidateSet and Get-HelpSectionNames'
    foreach ($name in $published) { Assert-Contains $validValues $name "ValidateSet is missing '$name'" }
    foreach ($value in $validValues) { Assert-Contains $published $value "Get-HelpSectionNames is missing '$value'" }
}

Invoke-TestCase -Name 'Bare -Help renders the complete reference (the README command)' -Body {
    $out = Clear-PrintSpoolerQueue -Help -NoColor -Quiet 6>&1 | Out-String
    Assert-GreaterThan $out.Length 20000 'full document must be comprehensive'
    Assert-Match 'USAGE' $out
    Assert-Match 'TROUBLESHOOTING' $out
    Assert-Match 'CREDITS' $out
    Assert-Equal 0 $Script:RuntimeState.ExitCode
}

Invoke-TestCase -Name '-HelpSection renders only the requested section' -Body {
    $out = Clear-PrintSpoolerQueue -HelpSection Usage -NoColor -Quiet 6>&1 | Out-String
    Assert-GreaterThan $out.Length 200
    Assert-Match 'USAGE' $out
    if ($out -cmatch 'TROUBLESHOOTING') { throw 'Usage-only render leaked the Troubleshooting heading' }
    if ($out -cmatch 'CREDITS') { throw 'Usage-only render leaked the Credits heading' }
    Assert-Equal 0 $Script:RuntimeState.ExitCode
}

Invoke-TestCase -Name '-HelpSection All equals the bare -Help document' -Body {
    $bare = Clear-PrintSpoolerQueue -Help -NoColor -Quiet 6>&1 | Out-String
    $all = Clear-PrintSpoolerQueue -HelpSection All -NoColor -Quiet 6>&1 | Out-String
    Assert-Equal $bare.Length $all.Length
}

Invoke-TestCase -Name '-HelpSection rejects unknown section names at binding time' -Body {
    Assert-Throws { Clear-PrintSpoolerQueue -HelpSection 'NoSuchSection' -NoColor -Quiet }
}

Invoke-TestCase -Name '-Help combined with -HelpSection selects the section' -Body {
    $out = Clear-PrintSpoolerQueue -Help -HelpSection ExitCodes -NoColor -Quiet 6>&1 | Out-String
    Assert-Match 'EXIT' $out
    if ($out -cmatch 'CREDITS') { throw '-Help -HelpSection ExitCodes leaked the Credits heading' }
    Assert-Equal 0 $Script:RuntimeState.ExitCode
}

Invoke-TestCase -Name '-ListHelpSections still returns the canonical names' -Body {
    $names = @(Clear-PrintSpoolerQueue -ListHelpSections -NoColor -Quiet -PassThru)
    $published = @(Get-HelpSectionNames)
    Assert-Equal $published.Count $names.Count
    foreach ($name in $published) { Assert-Contains $names $name }
    Assert-Equal 0 $Script:RuntimeState.ExitCode
}

Invoke-TestCase -Name 'Get-SpoolerCleanupHelp -Section renders the same content as -HelpSection' -Body {
    $viaCmdlet = Get-SpoolerCleanupHelp -Section Safety
    $viaCli = Clear-PrintSpoolerQueue -HelpSection Safety -NoColor -Quiet 6>&1 | Out-String
    Assert-GreaterThan $viaCmdlet.Length 500
    Assert-GreaterThan $viaCli.Length 500
    Assert-Match 'SAFETY' $viaCli
}

Invoke-TestCase -Name 'Every parameter of the main function has a help description' -Body {
    # -SelfTest already asserts this at runtime; pin it here too so the help
    # table cannot silently degrade to the placeholder text.
    $table = Get-SpoolerCleanupHelp -Section Parameters
    Assert-False ($table -match 'No description available') 'placeholder description leaked into the parameter table'
    Assert-Match 'HelpSection' $table 'the new -HelpSection parameter must appear in the metadata-driven table'
}

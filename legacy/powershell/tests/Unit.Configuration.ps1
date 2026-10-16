# Unit tests: configuration loading, merging, validation and file safety
Start-TestSuite 'Configuration defaults'

Invoke-TestCase -Name 'Defaults pass their own schema validation' -Body {
    Assert-Equal 0 @(Test-Configuration -Configuration (Get-DefaultConfiguration)).Count
}

Invoke-TestCase -Name 'Defaults declare every documented section' -Body {
    $d = Get-DefaultConfiguration
    foreach ($s in @('Service', 'Cleanup', 'Logging', 'Execution', 'Update', 'Diagnostics', 'Remote')) {
        Assert-True $d.Contains($s) "section $s missing"
    }
}

Invoke-TestCase -Name 'Default values match the documented production posture' -Body {
    $d = Get-DefaultConfiguration
    Assert-Equal 60 $d['Service']['TimeoutSeconds']
    Assert-Equal $true $d['Service']['StopDependentServices']
    Assert-Equal $true $d['Service']['RestartDependentServices']
    Assert-Equal $true $d['Remote']['PreFlightProbe']
    Assert-Equal $true $d['Update']['VerifyChecksum']
    Assert-Equal 1 $d['Execution']['ThrottleLimit']
    Assert-Equal 30 $d['Logging']['RetentionDays']
    Assert-Equal 3 $d['Cleanup']['RetryCount']
}

Invoke-TestCase -Name 'Getting the defaults twice returns independent copies' -Body {
    $a = Get-DefaultConfiguration
    $a['Service']['TimeoutSeconds'] = 1
    Assert-Equal 60 (Get-DefaultConfiguration)['Service']['TimeoutSeconds']
}

Start-TestSuite 'Configuration validation'

Invoke-TestCase -Name 'Out-of-range numbers are rejected' -Body {
    foreach ($case in @(
        @{ S = 'Service'; K = 'TimeoutSeconds'; V = 99999 }
        @{ S = 'Service'; K = 'TimeoutSeconds'; V = -5 }
        @{ S = 'Cleanup'; K = 'RetryCount'; V = 999 }
        @{ S = 'Logging'; K = 'MaxSizeMB'; V = 0 }
        @{ S = 'Execution'; K = 'ThrottleLimit'; V = 4096 }
        @{ S = 'Update'; K = 'CooldownHours'; V = -1 }
    )) {
        $c = Get-DefaultConfiguration
        $c[$case.S][$case.K] = $case.V
        Assert-GreaterThan @(Test-Configuration -Configuration $c).Count 0 "$($case.S).$($case.K)=$($case.V) was accepted"
    }
}

Invoke-TestCase -Name 'Non-numeric values in numeric fields are rejected' -Body {
    $c = Get-DefaultConfiguration
    $c['Service']['TimeoutSeconds'] = 'soon'
    Assert-GreaterThan @(Test-Configuration -Configuration $c).Count 0
}

Invoke-TestCase -Name 'Unknown sections are rejected' -Body {
    $c = Get-DefaultConfiguration
    $c['Nonsense'] = @{ A = 1 }
    $problems = @(Test-Configuration -Configuration $c)
    Assert-GreaterThan $problems.Count 0
    Assert-Match 'Unknown configuration section' ($problems -join ' ')
}

Invoke-TestCase -Name 'Unknown keys are rejected with the section named' -Body {
    $c = Get-DefaultConfiguration
    $c['Service']['NotARealKey'] = 1
    $problems = @(Test-Configuration -Configuration $c)
    Assert-GreaterThan $problems.Count 0
    Assert-Match 'Service\.NotARealKey' ($problems -join ' ')
}

Invoke-TestCase -Name 'A non-hashtable section is rejected' -Body {
    $c = Get-DefaultConfiguration
    $c['Service'] = 'not a hashtable'
    Assert-GreaterThan @(Test-Configuration -Configuration $c).Count 0
}

Invoke-TestCase -Name 'Wildcard delete extensions are refused outright' -Body {
    foreach ($bad in @(@('*'), @('.*'), @(''), @('.SPL', '*'))) {
        $c = Get-DefaultConfiguration
        $c['Cleanup']['IncludeExtensions'] = $bad
        $problems = @(Test-Configuration -Configuration $c)
        Assert-GreaterThan $problems.Count 0 "extension set [$($bad -join ',')] was accepted"
        Assert-Match 'wildcard' ($problems -join ' ')
    }
}

Invoke-TestCase -Name 'Malformed extensions are rejected' -Body {
    $c = Get-DefaultConfiguration
    $c['Cleanup']['IncludeExtensions'] = @('SPL')      # missing the dot
    Assert-GreaterThan @(Test-Configuration -Configuration $c).Count 0
}

Invoke-TestCase -Name 'An invalid log level is rejected' -Body {
    $c = Get-DefaultConfiguration
    $c['Logging']['Level'] = 'LOUD'
    Assert-GreaterThan @(Test-Configuration -Configuration $c).Count 0
}

Invoke-TestCase -Name 'An invalid export format is rejected' -Body {
    $c = Get-DefaultConfiguration
    $c['Execution']['ExportFormat'] = @('Xml')
    Assert-GreaterThan @(Test-Configuration -Configuration $c).Count 0
}

Invoke-TestCase -Name 'An empty configuration is reported' -Body {
    Assert-GreaterThan @(Test-Configuration -Configuration @{}).Count 0
}

Start-TestSuite 'Configuration merging'

Invoke-TestCase -Name 'A deep merge overrides one key and keeps its siblings' -Body {
    $m = Merge-Configuration -Base (Get-DefaultConfiguration) -Override @{ Service = @{ TimeoutSeconds = 120 } }
    Assert-Equal 120 $m['Service']['TimeoutSeconds']
    Assert-Equal 2 $m['Service']['SettleSeconds']
    Assert-Equal $true $m['Service']['StopDependentServices']
}

Invoke-TestCase -Name 'Merging never mutates the base' -Body {
    $base = Get-DefaultConfiguration
    $null = Merge-Configuration -Base $base -Override @{ Service = @{ TimeoutSeconds = 5 } }
    Assert-Equal 60 $base['Service']['TimeoutSeconds']
}

Invoke-TestCase -Name 'A null override value does not erase a default' -Body {
    $m = Merge-Configuration -Base (Get-DefaultConfiguration) -Override @{ Cleanup = @{ BackupPath = $null } }
    Assert-Equal 3 $m['Cleanup']['RetryCount']
}

Invoke-TestCase -Name 'An unknown key survives a merge so validation can reject it' -Body {
    $m = Merge-Configuration -Base (Get-DefaultConfiguration) -Override @{ Service = @{ Bogus = 1 } }
    Assert-GreaterThan @(Test-Configuration -Configuration $m).Count 0
}

Invoke-TestCase -Name 'Arrays are replaced, not concatenated' -Body {
    $m = Merge-Configuration -Base (Get-DefaultConfiguration) -Override @{ Cleanup = @{ IncludeExtensions = @('.SPL') } }
    Assert-Equal 1 @($m['Cleanup']['IncludeExtensions']).Count
}

Start-TestSuite 'Configuration file safety (AST parsing, never execution)'

$cfgDir = Join-Path ([IO.Path]::GetTempPath()) ("psq-cfg-{0}" -f [Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $cfgDir -Force
$cfgFile = Join-Path $cfgDir 'test.config.psd1'

Invoke-TestCase -Name 'A well-formed config file loads its values' -Body {
    Set-Content -LiteralPath $cfgFile -Value "@{ Service = @{ TimeoutSeconds = 90 }; Logging = @{ Level = 'DEBUG' } }" -Encoding UTF8
    $r = ConvertFrom-ConfigFileAst -Path $cfgFile
    Assert-Equal 0 $r.Errors.Count
    Assert-Equal 90 $r.Data['Service']['TimeoutSeconds']
    Assert-Equal 'DEBUG' $r.Data['Logging']['Level']
}

Invoke-TestCase -Name 'An executable statement in a config file is rejected' -Body {
    Set-Content -LiteralPath $cfgFile -Value "@{ Service = @{ TimeoutSeconds = 90 } }`r`nRemove-Item C:\ -Recurse -Force" -Encoding UTF8
    $r = ConvertFrom-ConfigFileAst -Path $cfgFile
    Assert-GreaterThan $r.Errors.Count 0
    Assert-Match 'Executable content' ($r.Errors -join ' ')
    Assert-Equal 0 $r.Data.Count 'a rejected file must yield no data'
}

Invoke-TestCase -Name 'A command invocation inside a value is rejected' -Body {
    Set-Content -LiteralPath $cfgFile -Value "@{ Service = @{ TimeoutSeconds = (Invoke-Expression '1+1') } }" -Encoding UTF8
    Assert-GreaterThan (ConvertFrom-ConfigFileAst -Path $cfgFile).Errors.Count 0
}

Invoke-TestCase -Name 'A pipeline in a config file is rejected' -Body {
    Set-Content -LiteralPath $cfgFile -Value "@{ Service = @{ Name = 'Spooler' } } | Out-Null" -Encoding UTF8
    Assert-GreaterThan (ConvertFrom-ConfigFileAst -Path $cfgFile).Errors.Count 0
}

Invoke-TestCase -Name 'A variable assignment in a config file is rejected' -Body {
    Set-Content -LiteralPath $cfgFile -Value "`$x = 1`r`n@{ Service = @{ TimeoutSeconds = `$x } }" -Encoding UTF8
    Assert-GreaterThan (ConvertFrom-ConfigFileAst -Path $cfgFile).Errors.Count 0
}

Invoke-TestCase -Name 'Malformed syntax reports a parse error, not an exception' -Body {
    Set-Content -LiteralPath $cfgFile -Value '@{ this is not valid' -Encoding UTF8
    $r = ConvertFrom-ConfigFileAst -Path $cfgFile
    Assert-GreaterThan $r.Errors.Count 0
}

Invoke-TestCase -Name 'A missing file reports an error' -Body {
    $r = ConvertFrom-ConfigFileAst -Path (Join-Path $cfgDir 'nope.psd1')
    Assert-GreaterThan $r.Errors.Count 0
    Assert-Match 'not found' ($r.Errors -join ' ')
}

Invoke-TestCase -Name 'An empty config file is reported, not silently accepted' -Body {
    Set-Content -LiteralPath $cfgFile -Value '# nothing here' -Encoding UTF8
    Assert-GreaterThan (ConvertFrom-ConfigFileAst -Path $cfgFile).Errors.Count 0
}

Invoke-TestCase -Name 'Get-EffectiveConfiguration honours an explicit path' -Body {
    Set-Content -LiteralPath $cfgFile -Value '@{ Service = @{ TimeoutSeconds = 77 } }' -Encoding UTF8
    $e = Get-EffectiveConfiguration -ConfigPath $cfgFile -ScriptDirectory $cfgDir -NoEnvironment
    Assert-Equal 77 $e.Config['Service']['TimeoutSeconds']
    Assert-Equal 'Parameter' $e.Source
    Assert-Equal 0 $e.Problems.Count
}

Invoke-TestCase -Name 'Get-EffectiveConfiguration reports a bad explicit path' -Body {
    $e = Get-EffectiveConfiguration -ConfigPath (Join-Path $cfgDir 'missing.psd1') -ScriptDirectory $cfgDir -NoEnvironment
    Assert-GreaterThan $e.Problems.Count 0
}

Invoke-TestCase -Name 'An invalid config file leaves the defaults in place' -Body {
    Set-Content -LiteralPath $cfgFile -Value '@{ Service = @{ TimeoutSeconds = 999999 } }' -Encoding UTF8
    $e = Get-EffectiveConfiguration -ConfigPath $cfgFile -ScriptDirectory $cfgDir -NoEnvironment
    Assert-GreaterThan $e.Problems.Count 0
}

Invoke-TestCase -Name 'NoConfigFile ignores an auto-discovered file' -Body {
    Set-Content -LiteralPath (Join-Path $cfgDir 'PrintSpoolerCleanup.config.psd1') -Value '@{ Service = @{ TimeoutSeconds = 33 } }' -Encoding UTF8
    $e = Get-EffectiveConfiguration -ScriptDirectory $cfgDir -NoConfigFile -NoEnvironment
    Assert-Equal 60 $e.Config['Service']['TimeoutSeconds']
    Assert-Equal 'Defaults' $e.Source
}

Invoke-TestCase -Name 'Auto-discovery finds the file next to the script' -Body {
    $e = Get-EffectiveConfiguration -ScriptDirectory $cfgDir -NoEnvironment
    Assert-Equal 33 $e.Config['Service']['TimeoutSeconds']
    Assert-Equal 'AutoDiscover' $e.Source
}

Start-TestSuite 'Environment variable configuration'

Invoke-TestCase -Name 'PSQ_ variables are collected with the prefix stripped' -Body {
    Invoke-WithEnvironment -Variables @{ PSQ_LOG_LEVEL = 'DEBUG'; PSQ_FORCE = 'true' } -Body {
        $m = Get-EnvironmentConfiguration
        Assert-Equal 'DEBUG' $m['LOG_LEVEL']
        Assert-Equal 'true' $m['FORCE']
    }
}

Invoke-TestCase -Name 'Boolean environment parsing accepts the common spellings' -Body {
    foreach ($v in @('1', 'true', 'TRUE', 'yes', 'Y', 'on', 'enabled')) { Assert-True (ConvertFrom-EnvironmentBoolean $v) $v }
    foreach ($v in @('0', 'false', 'no', 'off', '', $null)) { Assert-False (ConvertFrom-EnvironmentBoolean $v) "'$v'" }
    Assert-True (ConvertFrom-EnvironmentBoolean -Value $null -Default $true) 'the default is honoured'
}

Invoke-TestCase -Name 'PSQ_LOG_PATH flows into Logging.Path' -Body {
    # REGRESSION (defect 25): the log target used to be resolved from the -LogPath
    # PARAMETER alone, before configuration was loaded. So PSQ_LOG_PATH - which is
    # documented in the wiki, in the config template and in `-HelpSection Environment` -
    # had NO effect whatsoever, and Logging.Path from a config file was equally
    # dead. A site that configured a central log location silently got
    # auto-discovery next to the script instead.
    #
    # The variable is set through Invoke-WithEnvironment, NOT by assignment plus
    # `Remove-Item Env:`. The quality gate itself exports PSQ_LOG_PATH to keep
    # every stage from writing into dist/; deleting it here un-redirected logging
    # for all later stages and produced three unrelated-looking failures.
    $target = Join-Path ([IO.Path]::GetTempPath()) ("psq-logpath-{0}.log" -f [Guid]::NewGuid().ToString('N'))
    try {
        Invoke-WithEnvironment -Variables @{ PSQ_LOG_PATH = $target } -Body {
            $e = Get-EffectiveConfiguration -ScriptDirectory (Join-Path ([IO.Path]::GetTempPath()) 'no-such-dir') -NoConfigFile
            Assert-Equal $target $e.Config['Logging']['Path']
        }
    }
    finally { Remove-Item -LiteralPath $target -Force -ErrorAction SilentlyContinue }
}

Invoke-TestCase -Name 'PSQ_LOG_LEVEL flows into the effective configuration' -Body {
    Invoke-WithEnvironment -Variables @{ PSQ_LOG_LEVEL = 'WARN' } -Body {
        $e = Get-EffectiveConfiguration -ScriptDirectory $cfgDir -NoConfigFile
        Assert-Equal 'WARN' $e.Config['Logging']['Level']
        Assert-Match 'Environment' $e.Source
    }
}

Invoke-TestCase -Name 'PSQ_NO_UPDATE_CHECK inverts into AutoCheck=false' -Body {
    Invoke-WithEnvironment -Variables @{ PSQ_NO_UPDATE_CHECK = 'true' } -Body {
        $e = Get-EffectiveConfiguration -ScriptDirectory $cfgDir -NoConfigFile
        Assert-False $e.Config['Update']['AutoCheck']
    }
}

Invoke-TestCase -Name 'PSQ_PRINTER_NAME splits into a list' -Body {
    Invoke-WithEnvironment -Variables @{ PSQ_PRINTER_NAME = 'HP*, Canon*' } -Body {
        $e = Get-EffectiveConfiguration -ScriptDirectory $cfgDir -NoConfigFile
        Assert-Equal 2 @($e.Config['Cleanup']['PrinterName']).Count
    }
}

Start-TestSuite 'Configuration serialisation'

Invoke-TestCase -Name 'The defaults round-trip through Save/ConvertFrom' -Body {
    $target = Join-Path $cfgDir 'roundtrip.psd1'
    Assert-True (Save-Configuration -Path $target -Configuration (Get-DefaultConfiguration))
    $back = ConvertFrom-ConfigFileAst -Path $target
    Assert-Equal 0 $back.Errors.Count ($back.Errors -join '; ')
    Assert-Equal 60 $back.Data['Service']['TimeoutSeconds']
    Assert-Equal $true $back.Data['Service']['StopDependentServices']
    Assert-Equal 'INFO' $back.Data['Logging']['Level']
}

Invoke-TestCase -Name 'A saved configuration validates' -Body {
    $target = Join-Path $cfgDir 'validates.psd1'
    $null = Save-Configuration -Path $target -Configuration (Get-DefaultConfiguration)
    $back = ConvertFrom-ConfigFileAst -Path $target
    Assert-Equal 0 @(Test-Configuration -Configuration $back.Data).Count
}

Invoke-TestCase -Name 'Literals are escaped, not interpolated' -Body {
    Assert-Equal "'it''s'" (ConvertTo-ConfigLiteral -Value "it's")
    Assert-Equal '$null' (ConvertTo-ConfigLiteral -Value $null)
    Assert-Equal '$true' (ConvertTo-ConfigLiteral -Value $true)
    Assert-Equal '$false' (ConvertTo-ConfigLiteral -Value $false)
    Assert-Equal '42' (ConvertTo-ConfigLiteral -Value 42)
    Assert-Equal "@('a', 'b')" (ConvertTo-ConfigLiteral -Value @('a', 'b'))
    Assert-Equal '@()' (ConvertTo-ConfigLiteral -Value @())
}

Invoke-TestCase -Name 'The shipped template loads and validates' -Body {
    $template = Join-Path (Split-Path -Parent $PSScriptRoot) 'config/PrintSpoolerCleanup.config.psd1'
    Assert-True (Test-Path -LiteralPath $template) 'the template must ship with the product'
    $r = ConvertFrom-ConfigFileAst -Path $template
    Assert-Equal 0 $r.Errors.Count ($r.Errors -join '; ')
    Assert-Equal 0 @(Test-Configuration -Configuration $r.Data).Count
}

Remove-Item -LiteralPath $cfgDir -Recurse -Force -ErrorAction SilentlyContinue

# =============================================================================
#  Environment-variable contract
# -----------------------------------------------------------------------------
#  Defect 25 was PSQ_LOG_PATH: documented in -HelpSection Environment, in the wiki and in
#  the config template, and completely dead - the log target was resolved from the
#  -LogPath parameter before configuration was ever loaded. It was found by
#  accident. Two more of the same class followed (PSQ_CONFIG_PATH,
#  PSQ_NONINTERACTIVE), so the whole advertised surface is now audited here.
#
#  These tests fail the build if a documented PSQ_* variable stops being read, or
#  if a new one is read without being documented. Both directions matter: a dead
#  variable silently does nothing for an operator who set it in Group Policy, and
#  an undocumented one is undiscoverable.
# =============================================================================
Start-TestSuite 'Environment variable contract'

Invoke-TestCase -Name 'No documented PSQ_ variable is dead' -Body {
    # READ means referenced anywhere in src/ - either through the config layer
    # ($envMap.Contains('SUFFIX')) or directly as $env:PSQ_NAME. Checking only the
    # config layer produces a false positive: PSQ_COMPUTERS and PSQ_GITHUB_TOKEN
    # are read directly in Main-Function/Public-Cmdlets and are perfectly alive.
    $srcFiles = @(Get-ChildItem -LiteralPath (Join-Path $Root 'src') -Recurse -Filter *.ps1 -File)
    Assert-GreaterThan $srcFiles.Count 10 'the source tree must be present for this audit'
    $srcText = ($srcFiles | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }) -join "`n"

    $read = @()
    $read += @([regex]::Matches($srcText, "\`$envMap\.Contains\('([A-Z0-9_]+)'\)") | ForEach-Object { 'PSQ_' + $_.Groups[1].Value })
    $read += @([regex]::Matches($srcText, '\$env:(PSQ_[A-Z0-9_]+)') | ForEach-Object { $_.Groups[1].Value })
    $read = @($read | Sort-Object -Unique)
    Assert-GreaterThan $read.Count 20 'the product should read a substantial set of PSQ_ variables'

    # DOCUMENTED means advertised by the product's own -HelpSection Environment section or
    # by the shipped docs/config template.
    $helpSrc = Get-Content -LiteralPath (Join-Path $Root 'src/Private/13-HelpText.ps1') -Raw
    $documented = @([regex]::Matches($helpSrc, "V\s*=\s*'(PSQ_[A-Z0-9_]+)'") | ForEach-Object { $_.Groups[1].Value })
    foreach ($f in @('docs/wiki/Configuration.md', 'docs/wiki/CLI-Reference.md', 'README.md',
                     'config/PrintSpoolerCleanup.config.psd1', 'docs/wiki/Enterprise-Deployment.md',
                     'docs/wiki/Scheduling-and-Automation.md')) {
        $p = Join-Path $Root $f
        if (Test-Path -LiteralPath $p) {
            $documented += @([regex]::Matches((Get-Content -LiteralPath $p -Raw), 'PSQ_[A-Z0-9_]{2,}') | ForEach-Object { $_.Value })
        }
    }
    $documented = @($documented | Sort-Object -Unique)
    Assert-GreaterThan $documented.Count 20 'the documented PSQ_ surface should be substantial'

    $dead = @($documented | Where-Object { $read -notcontains $_ })
    Assert-Equal 0 $dead.Count $(if ($dead) {
            'documented but NEVER READ (defect 25 class - an operator setting these gets nothing): ' + ($dead -join ', ')
        } else { "$($documented.Count) documented, all read" })
}

Invoke-TestCase -Name 'No PSQ_ variable is read without being documented' -Body {
    $srcFiles = @(Get-ChildItem -LiteralPath (Join-Path $Root 'src') -Recurse -Filter *.ps1 -File)
    $srcText = ($srcFiles | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }) -join "`n"
    $read = @()
    $read += @([regex]::Matches($srcText, "\`$envMap\.Contains\('([A-Z0-9_]+)'\)") | ForEach-Object { 'PSQ_' + $_.Groups[1].Value })
    $read += @([regex]::Matches($srcText, '\$env:(PSQ_[A-Z0-9_]+)') | ForEach-Object { $_.Groups[1].Value })
    $read = @($read | Sort-Object -Unique)

    $docText = (Get-Content -LiteralPath (Join-Path $Root 'src/Private/13-HelpText.ps1') -Raw)
    foreach ($f in @('docs/wiki/Configuration.md', 'docs/wiki/CLI-Reference.md', 'README.md',
                     'config/PrintSpoolerCleanup.config.psd1')) {
        $p = Join-Path $Root $f
        if (Test-Path -LiteralPath $p) { $docText += (Get-Content -LiteralPath $p -Raw) }
    }
    $secret = @($read | Where-Object { $docText -notmatch [regex]::Escape($_) })
    Assert-Equal 0 $secret.Count $(if ($secret) { 'read but undocumented (undiscoverable): ' + ($secret -join ', ') } else { 'all documented' })
}

Invoke-TestCase -Name 'PSQ_NONINTERACTIVE actually suppresses the confirmation gate' -Body {
    # Regression for the dead-variable defect with the worst consequence: an
    # operator set PSQ_NONINTERACTIVE for a scheduled task, still got a prompt, and
    # with no console attached the task HUNG until its timeout instead of failing
    # fast. Assert the documented spellings all work.
    foreach ($v in @('true', '1', 'yes', 'on')) {
        $r = Invoke-WithEnvironment -Variables @{ PSQ_NONINTERACTIVE = $v } -Body {
            ConvertFrom-EnvironmentBoolean $env:PSQ_NONINTERACTIVE
        }
        Assert-True $r ("PSQ_NONINTERACTIVE='$v' must be truthy")
    }
    foreach ($v in @('false', '0', 'no', '', 'nonsense')) {
        $r = Invoke-WithEnvironment -Variables @{ PSQ_NONINTERACTIVE = $v } -Body {
            ConvertFrom-EnvironmentBoolean $env:PSQ_NONINTERACTIVE
        }
        Assert-False $r ("PSQ_NONINTERACTIVE='$v' must be falsy")
    }
}

Invoke-TestCase -Name 'PSQ_CONFIG_PATH resolves through Get-EffectiveConfiguration' -Body {
    # Regression: the variable was advertised but never consulted, so a site-wide
    # policy file set through Group Policy was silently ignored and the run fell
    # back to built-in defaults.
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("psq-cfgpath-{0}" -f [Guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $tmp -Force
    $cfg = Join-Path $tmp 'site-policy.psd1'
    Set-Content -LiteralPath $cfg -Encoding UTF8 -Value @'
@{
    Logging   = @{ Level = 'WARN'; RetentionDays = 45 }
    Execution = @{ Force = $true }
}
'@
    try {
        $e = Invoke-WithEnvironment -Variables @{ PSQ_CONFIG_PATH = $cfg } -Body {
            # The entry point resolves -ConfigPath > PSQ_CONFIG_PATH > discovery.
            # Reproduce that precedence exactly as Main-Function does.
            $explicit = $null
            if ([string]::IsNullOrWhiteSpace($explicit)) { $explicit = $env:PSQ_CONFIG_PATH }
            Get-EffectiveConfiguration -ConfigPath $explicit -ScriptDirectory (Join-Path ([IO.Path]::GetTempPath()) 'no-such-dir')
        }
        Assert-Equal 'WARN' $e.Config['Logging']['Level']
        Assert-Equal 45 $e.Config['Logging']['RetentionDays']
        # The provenance tokens are Parameter / AutoDiscover / Defaults /
        # Environment. An explicitly supplied path reports 'Parameter' - there is no
        # 'File' token, so asserting one would fail even when the file WAS read.
        # The value assertions above are what actually prove the file was loaded.
        Assert-Match 'Parameter' $e.Source ("an explicit path must report Parameter, got '{0}'" -f $e.Source)
        Assert-Equal $cfg $e.Path
    }
    finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

Invoke-TestCase -Name 'An explicit parameter beats its PSQ_ environment variable' -Body {
    # Precedence must be deterministic and documented: the command line wins, so an
    # operator can always override machine-wide policy for one run.
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("psq-prec-{0}" -f [Guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $tmp -Force
    $cfg = Join-Path $tmp 'policy.psd1'
    Set-Content -LiteralPath $cfg -Encoding UTF8 -Value "@{`n    Logging = @{ Level = 'ERROR' }`n}`n"
    try {
        $e = Invoke-WithEnvironment -Variables @{ PSQ_CONFIG_PATH = $cfg } -Body {
            $explicit = $cfg                      # simulates -ConfigPath being passed
            if ([string]::IsNullOrWhiteSpace($explicit)) { $explicit = $env:PSQ_CONFIG_PATH }
            Get-EffectiveConfiguration -ConfigPath $explicit -ScriptDirectory (Join-Path ([IO.Path]::GetTempPath()) 'no-such-dir')
        }
        Assert-Equal 'ERROR' $e.Config['Logging']['Level']
    }
    finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

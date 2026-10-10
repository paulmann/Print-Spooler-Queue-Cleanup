# ============================================================================
#  Configuration subsystem (finding A19)
# ----------------------------------------------------------------------------
#  v3 had no configuration story at all: changing the settle delay, the log
#  location or the default filters meant editing the script, which in turn broke
#  the next update. v4 introduces a declarative PSD1 config file with
#  auto-discovery, strict schema validation and a documented precedence chain:
#
#      1. Explicit command-line parameters          (highest)
#      2. PSQ_* environment variables
#      3. -ConfigPath <file>
#      4. Auto-discovered config file
#      5. Built-in defaults                          (lowest)
#
#  The config file is DATA ONLY (a hashtable literal). It is read with the
#  PowerShell AST and evaluated with SafeGetValue() - never with Invoke-Expression
#  on a file that may have come from a share - so a hostile config cannot execute
#  code.
# ============================================================================

function Get-DefaultConfiguration {
    <#
    .SYNOPSIS
        The complete built-in default configuration. This IS the schema reference.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()

    return @{
        Service = @{
            Name                     = 'Spooler'
            TimeoutSeconds           = 60
            SettleSeconds            = 2
            StopDependentServices    = $true
            RestartDependentServices = $true
            DependentServices        = @('PrintWorkflowUserSvc', 'Fax', 'PrintNotify')
            ForceStopFallback        = $true
        }
        Cleanup = @{
            IncludeExtensions        = @('.SPL', '.SHD')
            Recurse                  = $false
            Backup                   = $false
            BackupPath               = $null
            BackupKeepDays           = 30
            RetryCount               = 3
            RetryDelayMs             = 500
            WaitForQuiescenceSeconds = 0
            VerifyAfterRestart       = $true
            PurgeLegacyExtensions    = $false
            PrinterName              = @()
            ExcludePrinterName       = @()
            OlderThan                = $null
        }
        Logging = @{
            Level                  = 'INFO'
            Path                   = $null
            RetentionDays          = 30
            MaxSizeMB              = 10
            MaxFileCount           = 9
            JsonLog                = $false
            EventLog               = $false
            Timestamps             = $true
            SessionStampedFileName = $false
        }
        Execution = @{
            ThrottleLimit         = 1
            Force                 = $false
            NonInteractive        = $false
            ConfirmTimeoutSeconds = 0
            WatchdogMinutes       = 0
            SelfElevate           = $false
            NoColor               = $false
            Ascii                 = $false
            Quiet                 = $false
            NoProgress            = $false
            MachineReadableJson   = $false
            ExportFormat          = @()
            ExportPath            = $null
        }
        Update = @{
            Enabled           = $true
            AutoCheck         = $true
            CooldownHours     = 24
            Channel           = 'latest'
            IncludePrerelease = $false
            VerifyChecksum    = $true
            Repository        = 'paulmann/Print-Spooler-Queue-Cleanup'
            Branch            = 'main'
            Token             = $null
        }
        Diagnostics = @{
            IncludeEventLog    = $true
            EventLogHours      = 72
            EventLogMaxEntries = 200
            IncludeDriverStore = $true
            IncludePortMonitor = $true
            IncludeRegistry    = $true
        }
        Remote = @{
            PreFlightProbe      = $true
            ProbeTimeoutSeconds = 10
            SessionReuse        = $true
            Authentication      = 'Default'
            UseSsl              = $false
        }
    }
}

function Find-ConfigurationFile {
    <#
    .SYNOPSIS
        Auto-discovers PrintSpoolerCleanup.config.psd1 in the documented order.
    .OUTPUTS
        PSCustomObject { Path, Source, Exists } or $null.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyString()][string]$ExplicitPath,
        [AllowNull()][AllowEmptyString()][string]$ScriptDirectory,
        [string]$FileName = $Script:ConfigFileName
    )

    if (-not [string]::IsNullOrWhiteSpace($ExplicitPath)) {
        $full = [IO.Path]::GetFullPath($ExplicitPath)
        return [pscustomobject]@{ Path = $full; Source = 'Parameter'; Exists = (Test-Path -LiteralPath $full -PathType Leaf) }
    }

    $candidates = New-Object System.Collections.Generic.List[string]
    if ($ScriptDirectory) { $candidates.Add((Join-Path $ScriptDirectory $FileName)) }
    $candidates.Add((Join-Path (Get-Location).Path $FileName))
    if ($env:ProgramData) { $candidates.Add((Join-Path (Join-Path $env:ProgramData $Script:ProductFamily) $FileName)) }
    if ($env:LOCALAPPDATA) { $candidates.Add((Join-Path (Join-Path $env:LOCALAPPDATA $Script:ProductFamily) $FileName)) }

    foreach ($c in $candidates) {
        if ($c -and (Test-Path -LiteralPath $c -PathType Leaf)) {
            return [pscustomobject]@{ Path = $c; Source = 'AutoDiscover'; Exists = $true }
        }
    }
    return $null
}

function Convert-HashtableKeysToString {
    <#
    .SYNOPSIS
        Normalises nested dictionaries so lookups are always plain string keys.
    #>
    [CmdletBinding()]
    param([AllowNull()]$Hashtable)
    if ($null -eq $Hashtable) { return $null }
    if ($Hashtable -is [System.Collections.IDictionary]) {
        $out = @{}
        foreach ($k in @($Hashtable.Keys)) { $out[[string]$k] = Convert-HashtableKeysToString -Hashtable $Hashtable[$k] }
        return $out
    }
    if (($Hashtable -is [System.Collections.IEnumerable]) -and ($Hashtable -isnot [string])) {
        $list = New-Object System.Collections.Generic.List[object]
        foreach ($i in $Hashtable) { $list.Add((Convert-HashtableKeysToString -Hashtable $i)) }
        return $list.ToArray()
    }
    return $Hashtable
}

function ConvertFrom-ConfigFileAst {
    <#
    .SYNOPSIS
        Safely parses a .psd1 configuration file WITHOUT executing it.
    .DESCRIPTION
        Uses the PowerShell AST to locate a hashtable literal and evaluates only
        constant expressions via SafeGetValue(). Any command invocation, pipeline
        or assignment in the file is rejected. This makes it safe to accept a
        config that arrives from a file share or a Git repository.
    .OUTPUTS
        PSCustomObject { Data=[hashtable], Errors=[string[]] }
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Path)

    $errors = New-Object System.Collections.Generic.List[string]
    $result = [pscustomobject]@{ Data = @{}; Errors = @() }

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $result.Errors = @("Configuration file not found: $Path")
        return $result
    }
    if ((Get-Item -LiteralPath $Path).Length -gt 1MB) {
        $result.Errors = @("Configuration file exceeds 1 MB and was rejected: $Path")
        return $result
    }

    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) {
        foreach ($pe in $parseErrors) { $errors.Add("Parse error at line $($pe.Extent.StartLineNumber): $($pe.Message)") }
        $result.Errors = $errors.ToArray()
        return $result
    }
    if (-not $ast) {
        $errors.Add('The configuration file produced no syntax tree.')
        $result.Errors = $errors.ToArray()
        return $result
    }

    # Reject executable content: commands, multi-element pipelines, assignments.
    $badNodes = @($ast.FindAll({
                param($node)
                ($node -is [System.Management.Automation.Language.CommandAst]) -or
                ($node -is [System.Management.Automation.Language.PipelineAst] -and $node.PipelineElements.Count -gt 1) -or
                ($node -is [System.Management.Automation.Language.AssignmentStatementAst])
            }, $true))
    if ($badNodes.Count -gt 0) {
        foreach ($b in ($badNodes | Select-Object -First 5)) {
            $errors.Add("Executable content is not allowed in a config file (line $($b.Extent.StartLineNumber)): $($b.Extent.Text)")
        }
        $result.Errors = $errors.ToArray()
        return $result
    }

    $hashAst = @($ast.FindAll({
                param($node)
                $node -is [System.Management.Automation.Language.HashtableAst]
            }, $true) | Sort-Object { $_.Extent.StartLineNumber } | Select-Object -First 1)
    if ($hashAst.Count -eq 0 -or -not $hashAst[0]) {
        $errors.Add('No hashtable literal found in the configuration file.')
        $result.Errors = $errors.ToArray()
        return $result
    }

    try {
        $safe = $hashAst[0].SafeGetValue()
        if ($safe -isnot [System.Collections.IDictionary]) {
            $errors.Add('The top-level configuration value must be a hashtable.')
        }
        else {
            $result.Data = Convert-HashtableKeysToString -Hashtable $safe
        }
    }
    catch {
        $errors.Add("Configuration could not be evaluated: $($_.Exception.Message)")
    }
    $result.Errors = $errors.ToArray()
    return $result
}

function Test-ConfigSectionKeys {
    <#
    .SYNOPSIS
        Reports unknown or malformed keys inside one configuration section.
    .NOTES
        Deliberately a top-level private function: defining it nested inside
        Test-Configuration would leak a same-named function into script scope once
        the build inlines this file.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory)][AllowNull()][System.Collections.IDictionary]$Configuration,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Section,
        [Parameter(Mandatory)][AllowNull()][string[]]$Allowed
    )
    if (-not $Configuration -or -not $Configuration.Contains($Section)) { return @() }
    $sec = $Configuration[$Section]
    if ($sec -isnot [System.Collections.IDictionary]) { return @("Section '$Section' must be a hashtable.") }
    $bad = New-Object System.Collections.Generic.List[string]
    foreach ($k in @($sec.Keys)) {
        if (@($Allowed) -notcontains [string]$k) { $bad.Add("Unknown key '$Section.$k'.") }
    }
    return $bad.ToArray()
}

function Test-Configuration {
    <#
    .SYNOPSIS
        Schema and range validation. Returns human-readable problem strings.
    .DESCRIPTION
        A single typo in a machine-wide config file would otherwise silently
        disable backups or set a 0-second service timeout. Validation is loud:
        an invalid configuration aborts the run with exit code 3.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param([Parameter(Mandatory)][AllowNull()][System.Collections.IDictionary]$Configuration)

    $problems = New-Object System.Collections.Generic.List[string]

    # An empty hashtable is TRUTHY in PowerShell, so `-not $Configuration` never
    # fires for @{} - a config file that parsed to nothing would have been
    # reported as valid. Count the keys instead.
    if ($null -eq $Configuration -or @($Configuration.Keys).Count -eq 0) {
        return @('Configuration is empty: expected at least one section (Service, Cleanup, Logging, Execution, Update, Diagnostics, Remote).')
    }

    $defaults = Get-DefaultConfiguration
    $allowedSections = @($defaults.Keys)

    foreach ($k in @($Configuration.Keys)) {
        if ($allowedSections -notcontains [string]$k) {
            $problems.Add("Unknown configuration section '$k'. Allowed: $($allowedSections -join ', ').")
        }
    }

    foreach ($section in $allowedSections) {
        foreach ($msg in (Test-ConfigSectionKeys -Configuration $Configuration -Section $section -Allowed @($defaults[$section].Keys))) {
            $problems.Add($msg)
        }
    }

    # Range checks. Each entry is { Section, Key, Min, Max }.
    $checks = @(
        @{ S = 'Service'; K = 'TimeoutSeconds'; Min = 5; Max = 3600 }
        @{ S = 'Service'; K = 'SettleSeconds'; Min = 0; Max = 300 }
        @{ S = 'Cleanup'; K = 'RetryCount'; Min = 0; Max = 20 }
        @{ S = 'Cleanup'; K = 'RetryDelayMs'; Min = 0; Max = 60000 }
        @{ S = 'Cleanup'; K = 'BackupKeepDays'; Min = 0; Max = 3650 }
        @{ S = 'Cleanup'; K = 'WaitForQuiescenceSeconds'; Min = 0; Max = 600 }
        @{ S = 'Logging'; K = 'RetentionDays'; Min = 0; Max = 3650 }
        @{ S = 'Logging'; K = 'MaxSizeMB'; Min = 1; Max = 4096 }
        @{ S = 'Logging'; K = 'MaxFileCount'; Min = 1; Max = 99 }
        @{ S = 'Execution'; K = 'ThrottleLimit'; Min = 1; Max = 64 }
        @{ S = 'Execution'; K = 'WatchdogMinutes'; Min = 0; Max = 1440 }
        @{ S = 'Execution'; K = 'ConfirmTimeoutSeconds'; Min = 0; Max = 3600 }
        @{ S = 'Update'; K = 'CooldownHours'; Min = 0; Max = 8760 }
        @{ S = 'Diagnostics'; K = 'EventLogHours'; Min = 1; Max = 2160 }
        @{ S = 'Diagnostics'; K = 'EventLogMaxEntries'; Min = 1; Max = 5000 }
        @{ S = 'Remote'; K = 'ProbeTimeoutSeconds'; Min = 1; Max = 300 }
    )
    foreach ($c in $checks) {
        if (-not $Configuration.Contains($c.S)) { continue }
        $sec = $Configuration[$c.S]
        if ($sec -isnot [System.Collections.IDictionary]) { continue }
        if (-not $sec.Contains($c.K)) { continue }
        $raw = $sec[$c.K]
        if ($null -eq $raw) { continue }
        $num = 0.0
        if (-not [double]::TryParse([string]$raw, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$num)) {
            $problems.Add("$($c.S).$($c.K) must be a number (got '$raw').")
        }
        elseif ($num -lt $c.Min -or $num -gt $c.Max) {
            $problems.Add("$($c.S).$($c.K) must be between $($c.Min) and $($c.Max) (got $num).")
        }
    }

    if ($Configuration.Contains('Cleanup') -and $Configuration['Cleanup'] -is [System.Collections.IDictionary]) {
        $cleanup = $Configuration['Cleanup']
        if ($cleanup.Contains('IncludeExtensions')) {
            $exts = @($cleanup['IncludeExtensions'])
            foreach ($e in $exts) {
                $es = [string]$e
                if ($es -notmatch '^\.[A-Za-z0-9~]{1,6}$') {
                    $problems.Add("Cleanup.IncludeExtensions entry '$es' is not a valid extension (expected the form '.SPL').")
                }
            }
            # Hard safety rule: a wildcard here would turn the delete step into
            # "delete everything in the folder".
            if (($exts -contains '*') -or ($exts -contains '.*') -or ($exts -contains '')) {
                $problems.Add('Cleanup.IncludeExtensions must never contain a wildcard; the tool only deletes explicit spool extensions.')
            }
        }
    }

    if ($Configuration.Contains('Logging') -and $Configuration['Logging'] -is [System.Collections.IDictionary]) {
        $logging = $Configuration['Logging']
        if ($logging.Contains('Level')) {
            $lvl = [string]$logging['Level']
            if ($Script:LogLevelNames -notcontains $lvl.ToUpperInvariant()) {
                $problems.Add("Logging.Level '$lvl' is invalid. Allowed: $($Script:LogLevelNames -join ', ').")
            }
        }
    }

    if ($Configuration.Contains('Execution') -and $Configuration['Execution'] -is [System.Collections.IDictionary]) {
        $exec = $Configuration['Execution']
        if ($exec.Contains('ExportFormat')) {
            foreach ($fm in @($exec['ExportFormat'])) {
                if (@('Json', 'Csv', 'Html', 'Text') -notcontains ([string]$fm)) {
                    $problems.Add("Execution.ExportFormat '$fm' is invalid. Allowed: Json, Csv, Html, Text.")
                }
            }
        }
    }

    if ($Configuration.Contains('Update') -and $Configuration['Update'] -is [System.Collections.IDictionary]) {
        $upd = $Configuration['Update']
        if ($upd.Contains('Channel')) {
            $ch = [string]$upd['Channel']
            if (@('latest', 'stable', 'prerelease') -notcontains $ch.ToLowerInvariant()) {
                $problems.Add("Update.Channel '$ch' is invalid. Allowed: latest, stable, prerelease.")
            }
        }
    }

    return $problems.ToArray()
}

function Merge-Configuration {
    <#
    .SYNOPSIS
        Deep-merges $Override onto $Base. Neither input is mutated.
    .NOTES
        $null values in $Override are treated as "not set", so a config file cannot
        accidentally erase a built-in default by leaving a key empty.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)][AllowNull()][System.Collections.IDictionary]$Base,
        [AllowNull()][System.Collections.IDictionary]$Override
    )

    $merged = @{}
    if ($Base) { foreach ($k in @($Base.Keys)) { $merged[[string]$k] = $Base[$k] } }
    if (-not $Override) { return $merged }

    foreach ($k in @($Override.Keys)) {
        $key = [string]$k
        $value = $Override[$k]
        if ($null -eq $value) { continue }
        if ($merged.Contains($key) -and ($merged[$key] -is [System.Collections.IDictionary]) -and ($value -is [System.Collections.IDictionary])) {
            $merged[$key] = Merge-Configuration -Base (Convert-HashtableKeysToString -Hashtable $merged[$key]) -Override (Convert-HashtableKeysToString -Hashtable $value)
        }
        else {
            $merged[$key] = Convert-HashtableKeysToString -Hashtable $value
        }
    }
    return $merged
}

function Get-EffectiveConfiguration {
    <#
    .SYNOPSIS
        Builds the final runtime configuration from defaults + file + environment.
    .OUTPUTS
        PSCustomObject { Config, Source, Path, Problems }
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyString()][string]$ConfigPath,
        [AllowNull()][AllowEmptyString()][string]$ScriptDirectory,
        [switch]$NoConfigFile,
        [switch]$NoEnvironment
    )

    $defaults = Get-DefaultConfiguration
    $config = $defaults
    $source = 'Defaults'
    $usedPath = $null
    $problems = New-Object System.Collections.Generic.List[string]

    if (-not $NoConfigFile) {
        $found = Find-ConfigurationFile -ExplicitPath $ConfigPath -ScriptDirectory $ScriptDirectory
        if ($found -and $found.Exists) {
            $parsed = ConvertFrom-ConfigFileAst -Path $found.Path
            if ($parsed.Errors -and @($parsed.Errors).Count -gt 0) {
                foreach ($e in $parsed.Errors) { $problems.Add($e) }
            }
            else {
                $config = Merge-Configuration -Base $defaults -Override $parsed.Data
                $source = $found.Source
                $usedPath = $found.Path
            }
        }
        elseif ($found -and -not $found.Exists -and $ConfigPath) {
            $problems.Add("ConfigPath '$ConfigPath' does not exist.")
        }
    }

    if (-not $NoEnvironment) {
        $envMap = Get-EnvironmentConfiguration
        if (@($envMap.Keys).Count -gt 0) {
            $envCfg = @{}

            $logging = @{}
            if ($envMap.Contains('LOG_PATH')) { $logging['Path'] = $envMap['LOG_PATH'] }
            if ($envMap.Contains('LOG_LEVEL')) { $logging['Level'] = $envMap['LOG_LEVEL'] }
            if ($envMap.Contains('LOG_RETENTION_DAYS')) { $logging['RetentionDays'] = [int]$envMap['LOG_RETENTION_DAYS'] }
            if ($envMap.Contains('LOG_JSON')) { $logging['JsonLog'] = ConvertFrom-EnvironmentBoolean $envMap['LOG_JSON'] }
            if ($envMap.Contains('EVENT_LOG')) { $logging['EventLog'] = ConvertFrom-EnvironmentBoolean $envMap['EVENT_LOG'] }
            if (@($logging.Keys).Count -gt 0) { $envCfg['Logging'] = $logging }

            $service = @{}
            if ($envMap.Contains('SERVICE_TIMEOUT')) { $service['TimeoutSeconds'] = [int]$envMap['SERVICE_TIMEOUT'] }
            if ($envMap.Contains('SETTLE_SECONDS')) { $service['SettleSeconds'] = [int]$envMap['SETTLE_SECONDS'] }
            if (@($service.Keys).Count -gt 0) { $envCfg['Service'] = $service }

            $cleanup = @{}
            if ($envMap.Contains('RETRY_COUNT')) { $cleanup['RetryCount'] = [int]$envMap['RETRY_COUNT'] }
            if ($envMap.Contains('BACKUP_PATH')) { $cleanup['Backup'] = $true; $cleanup['BackupPath'] = $envMap['BACKUP_PATH'] }
            if ($envMap.Contains('OLDER_THAN')) { $cleanup['OlderThan'] = $envMap['OLDER_THAN'] }
            if ($envMap.Contains('PRINTER_NAME')) {
                # -split with a regex, NOT String.Split(',', ';'): that overload
                # binds to Split(char, int) and silently truncates instead of
                # splitting on both delimiters.
                $cleanup['PrinterName'] = @(($envMap['PRINTER_NAME'] -split '[,;]') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            }
            if (@($cleanup.Keys).Count -gt 0) { $envCfg['Cleanup'] = $cleanup }

            $execution = @{}
            if ($envMap.Contains('THROTTLE_LIMIT')) { $execution['ThrottleLimit'] = [int]$envMap['THROTTLE_LIMIT'] }
            if ($envMap.Contains('FORCE')) { $execution['Force'] = ConvertFrom-EnvironmentBoolean $envMap['FORCE'] }
            if ($envMap.Contains('QUIET')) { $execution['Quiet'] = ConvertFrom-EnvironmentBoolean $envMap['QUIET'] }
            if ($envMap.Contains('NO_COLOR')) { $execution['NoColor'] = ConvertFrom-EnvironmentBoolean $envMap['NO_COLOR'] }
            if ($envMap.Contains('ASCII')) { $execution['Ascii'] = ConvertFrom-EnvironmentBoolean $envMap['ASCII'] }
            if ($envMap.Contains('EXPORT_PATH')) { $execution['ExportPath'] = $envMap['EXPORT_PATH'] }
            if ($envMap.Contains('EXPORT_FORMAT')) { $execution['ExportFormat'] = @(($envMap['EXPORT_FORMAT'] -split '[,;]') | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
            if ($envMap.Contains('MACHINE_JSON')) { $execution['MachineReadableJson'] = ConvertFrom-EnvironmentBoolean $envMap['MACHINE_JSON'] }
            if (@($execution.Keys).Count -gt 0) { $envCfg['Execution'] = $execution }

            $update = @{}
            if ($envMap.Contains('UPDATE_COOLDOWN_HOURS')) { $update['CooldownHours'] = [int]$envMap['UPDATE_COOLDOWN_HOURS'] }
            if ($envMap.Contains('GITHUB_TOKEN')) { $update['Token'] = $envMap['GITHUB_TOKEN'] }
            if ($envMap.Contains('NO_UPDATE_CHECK')) { $update['AutoCheck'] = -not (ConvertFrom-EnvironmentBoolean $envMap['NO_UPDATE_CHECK']) }
            if ($envMap.Contains('UPDATE_CHANNEL')) { $update['Channel'] = $envMap['UPDATE_CHANNEL'] }
            if (@($update.Keys).Count -gt 0) { $envCfg['Update'] = $update }

            $config = Merge-Configuration -Base $config -Override $envCfg
            if ($source -eq 'Defaults') { $source = 'Environment' } else { $source = "$source+Environment" }
        }
    }

    foreach ($p in (Test-Configuration -Configuration $config)) { $problems.Add($p) }

    return [pscustomobject]@{
        Config   = $config
        Source   = $source
        Path     = $usedPath
        Problems = $problems.ToArray()
    }
}

function ConvertTo-ConfigLiteral {
    <#
    .SYNOPSIS
        Serialises a scalar, array or hashtable as a PowerShell literal.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()]$Value, [int]$Depth = 0)

    if ($null -eq $Value) { return '$null' }
    if ($Value -is [bool]) { return $(if ($Value) { '$true' } else { '$false' }) }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) { return [string]$Value }
    if ($Value -is [string]) { return "'{0}'" -f ($Value -replace "'", "''") }
    if ($Value -is [System.Collections.IDictionary]) {
        if ($Depth -ge 3) { return '@{}' }
        $inner = New-Object System.Collections.Generic.List[string]
        foreach ($k in (@($Value.Keys) | Sort-Object)) {
            $inner.Add(("{0} = {1}" -f $k, (ConvertTo-ConfigLiteral -Value $Value[$k] -Depth ($Depth + 1))))
        }
        return ('@{ ' + ($inner -join '; ') + ' }')
    }
    if ($Value -is [System.Collections.IEnumerable]) {
        $items = @($Value | ForEach-Object { ConvertTo-ConfigLiteral -Value $_ -Depth ($Depth + 1) })
        if ($items.Count -eq 0) { return '@()' }
        return ('@(' + ($items -join ', ') + ')')
    }
    return "'{0}'" -f ([string]$Value -replace "'", "''")
}

function Save-Configuration {
    <#
    .SYNOPSIS
        Writes a configuration hashtable back to a .psd1 file (-SaveConfig).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Path,
        [Parameter(Mandatory)][AllowNull()][System.Collections.IDictionary]$Configuration
    )
    try {
        $full = [IO.Path]::GetFullPath($Path)
        $dir = Split-Path -Parent $full
        if ($dir -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null
        }

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine('#')
        [void]$sb.AppendLine('# Print Spooler Queue Cleanup - configuration file')
        [void]$sb.AppendLine("# Generated by v$($Script:ScriptVersion) on $((Get-Date).ToString('u'))")
        [void]$sb.AppendLine('# Data only: this file is parsed as an AST and executable content is rejected.')
        [void]$sb.AppendLine('#')
        [void]$sb.AppendLine('@{')
        foreach ($section in (@($Configuration.Keys) | Sort-Object)) {
            [void]$sb.AppendLine("    $section = @{")
            $sec = $Configuration[$section]
            if ($sec -is [System.Collections.IDictionary]) {
                foreach ($key in (@($sec.Keys) | Sort-Object)) {
                    [void]$sb.AppendLine(("        {0,-28} = {1}" -f $key, (ConvertTo-ConfigLiteral -Value $sec[$key])))
                }
            }
            [void]$sb.AppendLine('    }')
        }
        [void]$sb.AppendLine('}')
        Set-Content -LiteralPath $full -Value $sb.ToString() -Encoding UTF8 -ErrorAction Stop
        return $true
    }
    catch {
        Write-LogError "Failed to save the configuration to '$Path': $($_.Exception.Message)" -Component 'Config'
        return $false
    }
}

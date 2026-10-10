# ============================================================================
#  Aggregation, console summary and export (JSON / CSV / HTML / Text)
# ----------------------------------------------------------------------------
#  Everything here is pure formatting: it consumes result objects and never
#  touches the Spooler. The HTML report is deliberately self-contained (inline
#  CSS, no CDN, no script) so it can be mailed out of an air-gapped network or
#  attached to a ticket without breaking - and so it cannot execute anything.
# ============================================================================

function Get-ResultPropertyValue {
    <#
    .SYNOPSIS
        StrictMode-safe property read used when aggregating heterogeneous results.

    .DESCRIPTION
        Reading `$item.SomeProperty` throws "The property 'SomeProperty' cannot be
        found on this object" under `Set-StrictMode -Version Latest` when the
        property is absent. Run results are deliberately heterogeneous - a Cleanup
        result carries JobsRemoved, a Diagnostic result does not - so aggregation
        must not assume a uniform shape.

        Going through PSObject.Properties returns $null for an absent property in
        every strict-mode setting, which also reproduces what
        `Measure-Object -Property` used to do silently (skip missing properties)
        without its other defect: returning $null for an empty pipeline, where
        `.Sum` then throws.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)][AllowNull()][object]$Item,
        [Parameter(Mandatory, Position = 1)][string]$Name
    )
    if ($null -eq $Item) { return $null }
    $prop = $Item.PSObject.Properties[$Name]
    if ($null -eq $prop) { return $null }
    return $prop.Value
}

function New-RunSummary {
    <#
    .SYNOPSIS
        Rolls per-computer results into one run-level record. Public schema.
    #>
    [CmdletBinding()]
    param(
        # AllowEmptyCollection matters: "every target was unreachable" is a real
        # and important outcome that must produce a Failed summary, not a
        # parameter-binding error.
        [Parameter(Mandatory)][AllowNull()][AllowEmptyCollection()][object[]]$Results,
        [Parameter(Mandatory)][AllowNull()][System.Collections.IDictionary]$Arguments,
        [AllowNull()][AllowEmptyString()][string]$LogPath,
        [AllowNull()][AllowEmptyString()][string]$ConfigSource,
        [AllowNull()]$Duration
    )

    $list = @($Results)
    # Statuses are read through Get-ResultPropertyValue rather than `$_.Status`.
    # Every object New-ResultObject produces does carry Status, but results also
    # arrive from ConvertFrom-PayloadResult (remote) and from caller-supplied
    # arrays, and under Set-StrictMode a single record missing the property would
    # abort the whole summary. Reading once into an array also keeps the seven
    # tallies below from re-walking the list seven times.
    $statuses = @($list | ForEach-Object { Get-ResultPropertyValue -Item $_ -Name 'Status' })
    $succeeded = @($statuses | Where-Object { $_ -eq 'Success' }).Count
    $withWarn = @($statuses | Where-Object { $_ -eq 'SuccessWithWarnings' }).Count
    $nothing = @($statuses | Where-Object { $_ -eq 'NothingToDo' }).Count
    $dryRun = @($statuses | Where-Object { $_ -eq 'DryRun' }).Count
    $diag = @($statuses | Where-Object { $_ -eq 'Diagnostic' }).Count
    $failed = @($statuses | Where-Object { $_ -in @('Failed', 'ServiceFailure') }).Count
    $skipped = @($statuses | Where-Object { $_ -in @('Skipped', 'Cancelled') }).Count

    $overall = 'Success'
    if ($list.Count -eq 0) { $overall = 'Failed' }
    elseif ($failed -gt 0 -and ($succeeded + $withWarn + $nothing + $dryRun + $diag) -gt 0) { $overall = 'PartialSuccess' }
    elseif ($failed -gt 0) { $overall = 'Failed' }
    elseif ($withWarn -gt 0) { $overall = 'SuccessWithWarnings' }
    elseif ($skipped -eq $list.Count) { $overall = 'Cancelled' }
    elseif ($dryRun -eq $list.Count) { $overall = 'DryRun' }
    elseif ($diag -eq $list.Count) { $overall = 'Diagnostic' }
    elseif ($nothing -eq $list.Count) { $overall = 'NothingToDo' }

    $elapsed = if ($null -eq $Duration) { [TimeSpan]::Zero } else { [TimeSpan]$Duration }

    # ---- StrictMode-safe aggregation -------------------------------------
    # Two distinct hazards, both invisible in lax mode:
    #
    #   1. Measure-Object returns $null (not an object) for an EMPTY pipeline, so
    #      `.Sum` throws "The property 'Sum' cannot be found on this object".
    #   2. Reading a property that does not exist - `$item.JobsFound` on a
    #      Diagnostic result - throws under Set-StrictMode -Version Latest,
    #      whereas Measure-Object -Property silently skipped missing properties.
    #
    # Result sets here are deliberately HETEROGENEOUS: a Cleanup result carries
    # JobsRemoved, a Diagnostic result does not. So property access goes through
    # PSObject.Properties, which returns $null for an absent property instead of
    # throwing, and empty collections are handled by the foreach simply not
    # executing. Correct in every mode, for every collection shape.
    $warnCount = 0; $errCount = 0
    foreach ($i in $list) {
        $warnCount += @(Get-ResultPropertyValue -Item $i -Name 'Warnings').Count
        $errCount += @(Get-ResultPropertyValue -Item $i -Name 'Errors').Count
    }

    $sumInt = {
        param($p)
        $t = 0
        foreach ($i in $list) {
            $v = Get-ResultPropertyValue -Item $i -Name $p
            if ($null -ne $v) { $t += [int]$v }
        }
        [int]$t
    }
    $sumLong = {
        param($p)
        $t = [long]0
        foreach ($i in $list) {
            $v = Get-ResultPropertyValue -Item $i -Name $p
            if ($null -ne $v) { $t += [long]$v }
        }
        [long]$t
    }

    return [pscustomobject]([ordered]@{
            Tool                = $Script:ScriptName
            ToolVersion         = $Script:ScriptVersion
            StartedUtc          = $Script:RuntimeState.RunStartedUtc.ToString('o')
            FinishedUtc         = [DateTime]::UtcNow.ToString('o')
            DurationMs          = [long]$elapsed.TotalMilliseconds
            Duration            = (Format-Duration -Duration $elapsed)
            OverallStatus       = $overall
            Mode                = $Arguments['Mode']
            DryRun              = [bool]$Arguments['DryRun']
            Operator            = (Get-CurrentUserContext).UserName
            HostComputer        = $env:COMPUTERNAME
            PowerShellVersion   = $PSVersionTable.PSVersion.ToString()
            PSEdition           = $PSVersionTable.PSEdition
            LogPath             = $LogPath
            ConfigSource        = $ConfigSource
            TargetsAttempted    = $list.Count
            TargetsSucceeded    = $succeeded
            TargetsWithWarnings = $withWarn
            TargetsNothingToDo  = $nothing
            TargetsDryRun       = $dryRun
            TargetsDiagnostic   = $diag
            TargetsFailed       = $failed
            TargetsSkipped      = $skipped
            JobsFound           = (& $sumInt 'JobsFound')
            JobsSelected        = (& $sumInt 'JobsSelected')
            JobsRemoved         = (& $sumInt 'JobsRemoved')
            FilesFound          = (& $sumInt 'FilesFound')
            FilesDeleted        = (& $sumInt 'FilesDeleted')
            FilesFailed         = (& $sumInt 'FilesFailed')
            FilesLocked         = (& $sumInt 'FilesLocked')
            ResurrectedFiles    = (& $sumInt 'ResurrectedFiles')
            BytesFound          = (& $sumLong 'BytesFound')
            BytesDeleted        = (& $sumLong 'BytesDeleted')
            Warnings            = [int]($warnCount)
            Errors              = [int]($errCount)
            BackupsCreated      = @($list | ForEach-Object { Get-ResultPropertyValue -Item $_ -Name 'BackupPath' } | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
            ExitCode            = $null
            Results             = $list
        })
}

function Get-ExitCodeForSummary {
    <#
    .SYNOPSIS
        Maps a run summary onto the public exit-code contract.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param([Parameter(Mandatory)][AllowNull()]$Summary)

    switch ($Summary.OverallStatus) {
        'Success' { return $Script:ExitCode.Success }
        'NothingToDo' { return $Script:ExitCode.Success }
        'DryRun' { return $Script:ExitCode.Success }
        'Diagnostic' { return $Script:ExitCode.DiagnosticsOnly }
        'SuccessWithWarnings' { return $Script:ExitCode.PartialSuccess }
        'PartialSuccess' { return $Script:ExitCode.PartialSuccess }
        'Cancelled' { return $Script:ExitCode.Cancelled }
        default { return $Script:ExitCode.Failed }
    }
}

function Write-RunSummaryToConsole {
    <#
    .SYNOPSIS
        The EXECUTION SUMMARY block operators actually read.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowNull()]$Summary)

    if ($Script:RuntimeState.Quiet -or $Script:RuntimeState.MachineReadable) { return }
    if (-not $Summary) { return }

    $bar = Get-Bar -Character '='
    $hdr = if ($Script:RuntimeState.NoColor) { $null } else { $Script:RuntimeState.Colors['Header'] }
    Write-Host ''
    if ($null -ne $hdr) {
        Write-Host $bar -ForegroundColor $hdr
        Write-Host '    EXECUTION SUMMARY' -ForegroundColor $hdr
        Write-Host $bar -ForegroundColor $hdr
    }
    else {
        Write-Host $bar
        Write-Host '    EXECUTION SUMMARY'
        Write-Host $bar
    }

    $statusColor = switch ($Summary.OverallStatus) {
        'Success' { 'Green' }
        'NothingToDo' { 'Green' }
        'DryRun' { 'Cyan' }
        'Diagnostic' { 'Cyan' }
        'SuccessWithWarnings' { 'Yellow' }
        'PartialSuccess' { 'Yellow' }
        'Cancelled' { 'Yellow' }
        default { 'Red' }
    }

    Write-ConsoleKeyValue -Label 'Overall result' -Value $Summary.OverallStatus -ValueColor $statusColor
    Write-ConsoleKeyValue -Label 'Mode' -Value $Summary.Mode
    Write-ConsoleKeyValue -Label 'Duration' -Value $Summary.Duration
    Write-ConsoleKeyValue -Label 'Targets' -Value ("{0} attempted / {1} succeeded / {2} warned / {3} failed" -f $Summary.TargetsAttempted, $Summary.TargetsSucceeded, $Summary.TargetsWithWarnings, $Summary.TargetsFailed)
    Write-ConsoleKeyValue -Label 'Print jobs found' -Value $Summary.JobsFound -ValueColor $(if ($Summary.JobsFound -gt 0) { 'White' } else { 'DarkGray' })
    Write-ConsoleKeyValue -Label 'Jobs removed' -Value $Summary.JobsRemoved -ValueColor $(if ($Summary.JobsRemoved -gt 0) { 'Green' } else { 'DarkGray' })
    Write-ConsoleKeyValue -Label 'Files deleted' -Value $Summary.FilesDeleted -ValueColor $(if ($Summary.FilesDeleted -gt 0) { 'Green' } else { 'DarkGray' })
    Write-ConsoleKeyValue -Label 'Space reclaimed' -Value (Format-ByteSize $Summary.BytesDeleted) -ValueColor $(if ($Summary.BytesDeleted -gt 0) { 'Green' } else { 'DarkGray' })
    if ($Summary.FilesLocked -gt 0) { Write-ConsoleKeyValue -Label 'Files still locked' -Value $Summary.FilesLocked -ValueColor 'Yellow' }
    if ($Summary.ResurrectedFiles -gt 0) { Write-ConsoleKeyValue -Label 'Files resurrected' -Value $Summary.ResurrectedFiles -ValueColor 'Yellow' }
    $weColor = if ($Summary.Errors -gt 0) { 'Red' } elseif ($Summary.Warnings -gt 0) { 'Yellow' } else { 'Green' }
    Write-ConsoleKeyValue -Label 'Warnings / errors' -Value ("{0} / {1}" -f $Summary.Warnings, $Summary.Errors) -ValueColor $weColor
    if ($Summary.BackupsCreated -and @($Summary.BackupsCreated).Count -gt 0) {
        Write-ConsoleKeyValue -Label 'Backup(s)' -Value (@($Summary.BackupsCreated) -join '; ')
    }
    if ($Summary.LogPath) { Write-ConsoleKeyValue -Label 'Log file' -Value $Summary.LogPath }
    Write-ConsoleKeyValue -Label 'Exit code' -Value $Summary.ExitCode -ValueColor $(if ($Summary.ExitCode -eq 0) { 'Green' } else { 'Yellow' })

    $rows = @($Summary.Results)
    if ($rows.Count -gt 0 -and $Summary.Mode -ne 'Diagnostic') {
        Write-Host ''
        Write-ConsoleSection -Title 'PER-COMPUTER RESULTS'
        $table = foreach ($r in $rows) {
            [pscustomobject]@{
                Computer       = $r.ComputerName
                Status         = $r.Status
                Jobs           = $r.JobsRemoved
                Deleted        = $r.FilesDeleted
                Locked         = $r.FilesLocked
                Reclaimed      = (Format-ByteSize $r.BytesDeleted)
                ServiceRunning = $(if ($r.ServiceIsRunning) { 'yes' } elseif ($null -eq $r.ServiceIsRunning) { '?' } else { 'NO' })
                Seconds        = [Math]::Round($r.DurationMs / 1000, 1)
            }
        }
        $table | Write-ConsoleTable -Property @('Computer', 'Status', 'Jobs', 'Deleted', 'Locked', 'Reclaimed', 'ServiceRunning', 'Seconds')
    }

    $allErrors = @()
    foreach ($r in $rows) { foreach ($e in @($r.Errors)) { $allErrors += [pscustomobject]@{ Computer = $r.ComputerName; Text = $e } } }
    if ($allErrors.Count -gt 0) {
        Write-Host ''
        Write-ConsoleSection -Title 'ERRORS' -Character '='
        foreach ($e in ($allErrors | Select-Object -First 12)) {
            $line = "    [{0}] {1}" -f $e.Computer, $e.Text
            if ($Script:RuntimeState.NoColor) { Write-Host $line } else { Write-Host $line -ForegroundColor 'Red' }
        }
        if ($allErrors.Count -gt 12) { Write-Host ("    ... and {0} more (see the log file)" -f ($allErrors.Count - 12)) }
    }

    $allWarnings = @()
    foreach ($r in $rows) { foreach ($w in @($r.Warnings)) { $allWarnings += [pscustomobject]@{ Computer = $r.ComputerName; Text = $w } } }
    if ($allWarnings.Count -gt 0) {
        Write-Host ''
        Write-ConsoleSection -Title 'WARNINGS' -Character '='
        foreach ($w in ($allWarnings | Select-Object -First 10)) {
            $line = "    [{0}] {1}" -f $w.Computer, $w.Text
            if ($Script:RuntimeState.NoColor) { Write-Host $line } else { Write-Host $line -ForegroundColor 'Yellow' }
        }
        if ($allWarnings.Count -gt 10) { Write-Host ("    ... and {0} more (see the log file)" -f ($allWarnings.Count - 10)) }
    }

    if ($null -ne $hdr) { Write-Host $bar -ForegroundColor $hdr } else { Write-Host $bar }
}

function Write-DiagnosticReport {
    <#
    .SYNOPSIS
        Human-readable rendering of -Diagnostic output (finding A5).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowNull()]$Summary)

    if ($Script:RuntimeState.Quiet -or $Script:RuntimeState.MachineReadable) { return }
    if (-not $Summary) { return }

    foreach ($r in @($Summary.Results)) {
        Write-Host ''
        Write-ConsoleSection -Title "DIAGNOSTIC REPORT - $($r.ComputerName)" -Character '='
        $d = $r.Diagnostics
        if (-not $d) { Write-Host '    (no diagnostics returned)'; continue }

        if ($d.Contains('OS')) {
            Write-ConsoleKeyValue -Label 'Operating system' -Value ("{0} ({1}, build {2})" -f $d['OS']['Caption'], $d['OS']['Version'], $d['OS']['Build'])
            if ($d['OS']['LastBoot']) { Write-ConsoleKeyValue -Label 'Last boot' -Value $d['OS']['LastBoot'] }
            Write-ConsoleKeyValue -Label 'Free memory' -Value ("{0} MB" -f $d['OS']['FreeMemoryMB'])
        }

        Write-Host ''
        Write-ConsoleSection -Title 'Print-related services'
        if ($d.Contains('Services')) { @($d['Services']) | Write-ConsoleTable -Property @('Name', 'State', 'StartMode', 'ProcessId', 'Exists') }

        if ($d.Contains('SpoolerProcess') -and @($d['SpoolerProcess']).Count -gt 0) {
            Write-Host ''
            Write-ConsoleSection -Title 'Spooler processes'
            @($d['SpoolerProcess']) | Write-ConsoleTable -Property @('Name', 'Id', 'CPUSeconds', 'WorkingSetMB', 'Threads', 'Responding')
        }

        if ($d.Contains('SpoolDirectory')) {
            Write-Host ''
            Write-ConsoleSection -Title "Spool directory - $($r.SpoolDirectory)"
            $sd = $d['SpoolDirectory']
            Write-ConsoleKeyValue -Label 'Files' -Value $sd['FileCount']
            Write-ConsoleKeyValue -Label 'Total size' -Value (Format-ByteSize ([long]$sd['TotalBytes']))
            Write-ConsoleKeyValue -Label 'Owner' -Value $sd['AclOwner']
            if ($sd['OldestFile']) {
                # AgeHours is fractional; format it invariantly so a host with a comma
                # decimal separator still reads "12.5 h old" rather than "12,5 h old".
                $ageText = ([double]$sd['OldestFile']['AgeHours']).ToString([Globalization.CultureInfo]::InvariantCulture)
                $ageColor = if ([double]$sd['OldestFile']['AgeHours'] -gt 24) { 'Red' } else { 'White' }
                Write-ConsoleKeyValue -Label 'Oldest file' -Value ("{0} ({1} h old)" -f $sd['OldestFile']['Name'], $ageText) -ValueColor $ageColor
            }
            if ($sd['ByExtension']) {
                Write-Host ''
                $ext = foreach ($x in @($sd['ByExtension'])) {
                    [pscustomobject]@{ Extension = $x['Ext']; Count = $x['Count']; Bytes = (Format-ByteSize ([long]$x['Bytes'])) }
                }
                $ext | Write-ConsoleTable -Property @('Extension', 'Count', 'Bytes')
            }
        }

        $jobs = @($r.Jobs)
        Write-Host ''
        if ($jobs.Count -gt 0) {
            Write-ConsoleSection -Title "Queued jobs ($($jobs.Count))"
            $jobRows = foreach ($j in ($jobs | Select-Object -First 40)) {
                [pscustomobject]@{
                    JobId     = $j.JobId
                    Printer   = $j.PrinterName
                    Document  = $j.DocumentName
                    Owner     = $j.UserName
                    Size      = (Format-ByteSize ([long]$j.SizeBytes))
                    Pages     = $j.TotalPages
                    Status    = $j.Status
                    Submitted = if ($j.SubmittedUtc) { ([DateTime]$j.SubmittedUtc).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { 'n/a' }
                    Source    = $j.Source
                }
            }
            $jobRows | Write-ConsoleTable -Property @('JobId', 'Printer', 'Document', 'Owner', 'Size', 'Pages', 'Status', 'Submitted', 'Source')
        }
        else {
            Write-ConsoleSection -Title 'Queued jobs'
            Write-Host '    No queued print jobs were found.'
        }

        if ($d.Contains('Printers') -and @($d['Printers']).Count -gt 0) {
            Write-Host ''
            Write-ConsoleSection -Title "Printers ($(@($d['Printers']).Count))"
            $printerRows = foreach ($p in (@($d['Printers']) | Select-Object -First 40)) {
                [pscustomobject]@{
                    Name    = $p.Name
                    State   = $p.PrinterState
                    Status  = $p.PrinterStatus
                    Local   = $(if ($p.Local) { 'yes' } else { 'no' })
                    Shared  = $(if ($p.Shared) { 'yes' } else { 'no' })
                    Offline = $(if ($p.WorkOffline) { 'YES' } else { 'no' })
                    Port    = $p.PortName
                    Driver  = $p.DriverName
                }
            }
            $printerRows | Write-ConsoleTable -Property @('Name', 'State', 'Status', 'Local', 'Shared', 'Offline', 'Port', 'Driver')
        }

        if ($d.Contains('EventLog') -and @($d['EventLog']).Count -gt 0) {
            Write-Host ''
            Write-ConsoleSection -Title 'Recent Service Control Manager events (Spooler/Print)'
            foreach ($ev in (@($d['EventLog']) | Select-Object -First 15)) {
                $msg = [string]$ev.Message
                if ($msg.Length -gt 150) { $msg = $msg.Substring(0, 147) + '...' }
                Write-Host ("    {0}  [{1}] {2}" -f ([DateTime]$ev.TimeCreated).ToString('yyyy-MM-dd HH:mm:ss'), $ev.LevelDisplayName, $msg)
            }
        }

        if ($d.Contains('PrintEventLog') -and @($d['PrintEventLog']).Count -gt 0) {
            Write-Host ''
            Write-ConsoleSection -Title 'Microsoft-Windows-PrintService/Admin events'
            foreach ($ev in (@($d['PrintEventLog']) | Select-Object -First 15)) {
                $msg = [string]$ev.Message
                if ($msg.Length -gt 150) { $msg = $msg.Substring(0, 147) + '...' }
                Write-Host ("    {0}  [{1}] id={2} {3}" -f ([DateTime]$ev.TimeCreated).ToString('yyyy-MM-dd HH:mm:ss'), $ev.LevelDisplayName, $ev.Id, $msg)
            }
        }
    }
}

function ConvertTo-CamelCaseName {
    <#
    .SYNOPSIS
        PascalCase -> camelCase, handling leading acronyms the way the published
        report schema spells them.

    .DESCRIPTION
        A single mechanical rule, with no per-field special cases, so the emitter
        and docs/Report-Schema.md cannot drift apart field by field:

            ToolVersion         -> toolVersion          (run of 1)
            PSEdition           -> psEdition            (run of 3, then lowercase)
            ID                  -> id                   (all uppercase)
            alreadyCamel        -> alreadyCamel         (untouched)

        For a leading run of N uppercase letters followed by a lowercase letter,
        the first N-1 are lowercased and the last one starts the next word. That
        is what produces `psEdition` rather than `pSEdition`.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][AllowNull()][AllowEmptyString()][string]$Name)

    if ([string]::IsNullOrEmpty($Name)) { return $Name }
    if (-not [char]::IsUpper($Name[0])) { return $Name }

    $run = 0
    while ($run -lt $Name.Length -and [char]::IsUpper($Name[$run])) { $run++ }

    if ($run -eq 1) { return ($Name.Substring(0, 1).ToLowerInvariant() + $Name.Substring(1)) }
    if ($run -lt $Name.Length) { return ($Name.Substring(0, $run - 1).ToLowerInvariant() + $Name.Substring($run - 1)) }
    return $Name.ToLowerInvariant()
}

function ConvertTo-CamelCaseDeep {
    <#
    .SYNOPSIS
        Recursively rebuilds a report graph with camelCase keys, values untouched.

    .DESCRIPTION
        The JSON report is a PUBLISHED, normative contract (docs/Report-Schema.md)
        that RMM tooling, jq and Python consume. Those consumers are
        CASE-SENSITIVE. Serialising the PowerShell object directly emitted
        PascalCase keys inside a camelCase envelope - so every documented field
        name, and both the jq and Python examples in the schema document, resolved
        to nothing.

        This went undetected because PowerShell property access is
        case-insensitive: `$obj.summary.Tool` and `$obj.summary.tool` both work,
        so a test written in PowerShell cannot see the difference. Only the raw
        wire text reveals it.

        Leaf values are passed through untouched so ConvertTo-Json still renders
        numbers, booleans and dates itself. Depth is bounded because report graphs
        are shallow and an accidental cycle must not hang a run.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowNull()]$Value, [int]$Depth = 12)

    if ($null -eq $Value) { return $null }
    if ($Depth -le 0) { return $Value }

    # Leaves: strings and every value type (bool, int, long, double, DateTime,
    # TimeSpan, Guid, enums, SwitchParameter). Let ConvertTo-Json format these.
    if ($Value -is [string]) { return $Value }
    if ($Value -is [ValueType]) { return $Value }

    if ($Value -is [System.Collections.IDictionary]) {
        $out = [ordered]@{}
        foreach ($k in @($Value.Keys)) {
            $out[(ConvertTo-CamelCaseName -Name ([string]$k))] = ConvertTo-CamelCaseDeep -Value $Value[$k] -Depth ($Depth - 1)
        }
        return $out
    }

    if ($Value -is [System.Collections.IEnumerable]) {
        $list = New-Object System.Collections.Generic.List[object]
        foreach ($item in $Value) { $list.Add((ConvertTo-CamelCaseDeep -Value $item -Depth ($Depth - 1))) }
        # The unary comma matters: without it a SINGLE-element array unrolls to a
        # scalar on assignment and the JSON emits an object instead of an array.
        return , ($list.ToArray())
    }

    $props = @($Value.PSObject.Properties)
    if ($props.Count -gt 0) {
        $out = [ordered]@{}
        foreach ($pr in $props) {
            $out[(ConvertTo-CamelCaseName -Name $pr.Name)] = ConvertTo-CamelCaseDeep -Value $pr.Value -Depth ($Depth - 1)
        }
        return $out
    }
    return $Value
}

function ConvertTo-JsonReport {
    <#
    .SYNOPSIS
        Full-fidelity JSON. This is the machine contract consumed by RMM tools.
    .DESCRIPTION
        Keys are camelCase to match the normative schema document. See
        ConvertTo-CamelCaseDeep for why this cannot be left to ConvertTo-Json.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][AllowNull()]$Summary, [int]$Depth = 8)

    $payload = [ordered]@{
        schemaVersion = '4.0'
        summary       = (ConvertTo-CamelCaseDeep -Value $Summary)
    }
    return ($payload | ConvertTo-Json -Depth $Depth)
}

function ConvertTo-CsvReport {
    <#
    .SYNOPSIS
        One row per target computer, for spreadsheets and BI.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][AllowNull()]$Summary)

    $rows = @(foreach ($r in @($Summary.Results)) {
            [pscustomobject]([ordered]@{
                    RunStartedUtc     = $Summary.StartedUtc
                    ToolVersion       = $Summary.ToolVersion
                    ComputerName      = $r.ComputerName
                    IsLocal           = $r.IsLocal
                    Status            = $r.Status
                    SpoolDirectory    = $r.SpoolDirectory
                    ServiceWasRunning = $r.ServiceWasRunning
                    ServiceIsRunning  = $r.ServiceIsRunning
                    JobsFound         = $r.JobsFound
                    JobsSelected      = $r.JobsSelected
                    JobsRemoved       = $r.JobsRemoved
                    FilesFound        = $r.FilesFound
                    FilesDeleted      = $r.FilesDeleted
                    FilesFailed       = $r.FilesFailed
                    FilesLocked       = $r.FilesLocked
                    BytesFound        = $r.BytesFound
                    BytesDeleted      = $r.BytesDeleted
                    FilesAfter        = $r.FilesAfter
                    ResurrectedFiles  = $r.ResurrectedFiles
                    ServiceStopMs     = $r.ServiceStopMs
                    ServiceStartMs    = $r.ServiceStartMs
                    DurationMs        = $r.DurationMs
                    BackupPath        = $r.BackupPath
                    BackupHash        = $r.BackupHash
                    WarningCount      = @($r.Warnings).Count
                    ErrorCount        = @($r.Errors).Count
                    FirstError        = (@($r.Errors) | Select-Object -First 1)
                })
        })
    if ($rows.Count -eq 0) { return '' }
    return (($rows | ConvertTo-Csv -NoTypeInformation) -join [Environment]::NewLine)
}

function ConvertTo-TextReport {
    <#
    .SYNOPSIS
        Plain-text report for tickets, pasting into an email, or `more`.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][AllowNull()]$Summary)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('=' * 78)
    [void]$sb.AppendLine(("{0} v{1} - run report" -f $Script:ScriptName, $Script:ScriptVersion))
    [void]$sb.AppendLine('=' * 78)
    [void]$sb.AppendLine(("Started (UTC) : {0}" -f $Summary.StartedUtc))
    [void]$sb.AppendLine(("Finished(UTC) : {0}" -f $Summary.FinishedUtc))
    [void]$sb.AppendLine(("Duration      : {0}" -f $Summary.Duration))
    [void]$sb.AppendLine(("Overall       : {0}" -f $Summary.OverallStatus))
    [void]$sb.AppendLine(("Mode          : {0}" -f $Summary.Mode))
    [void]$sb.AppendLine(("Operator      : {0}" -f $Summary.Operator))
    [void]$sb.AppendLine(("Targets       : {0} attempted, {1} succeeded, {2} warned, {3} failed" -f $Summary.TargetsAttempted, $Summary.TargetsSucceeded, $Summary.TargetsWithWarnings, $Summary.TargetsFailed))
    [void]$sb.AppendLine(("Jobs          : {0} found, {1} selected, {2} removed" -f $Summary.JobsFound, $Summary.JobsSelected, $Summary.JobsRemoved))
    [void]$sb.AppendLine(("Files         : {0} found, {1} deleted, {2} failed, {3} locked" -f $Summary.FilesFound, $Summary.FilesDeleted, $Summary.FilesFailed, $Summary.FilesLocked))
    [void]$sb.AppendLine(("Bytes         : {0} found, {1} reclaimed" -f (Format-ByteSize $Summary.BytesFound), (Format-ByteSize $Summary.BytesDeleted)))
    [void]$sb.AppendLine(("Exit code     : {0}" -f $Summary.ExitCode))
    [void]$sb.AppendLine('')

    foreach ($r in @($Summary.Results)) {
        [void]$sb.AppendLine(('-' * 78))
        [void]$sb.AppendLine(("{0}  [{1}]" -f $r.ComputerName, $r.Status))
        [void]$sb.AppendLine(("    Spool directory : {0} ({1})" -f $r.SpoolDirectory, $r.SpoolDirectorySource))
        [void]$sb.AppendLine(("    Service         : was {0}, now {1}" -f $(if ($r.ServiceWasRunning) { 'Running' } else { 'Stopped' }), $(if ($r.ServiceIsRunning) { 'Running' } else { 'NOT RUNNING' })))
        [void]$sb.AppendLine(("    Jobs            : found {0}, selected {1}, removed {2}" -f $r.JobsFound, $r.JobsSelected, $r.JobsRemoved))
        [void]$sb.AppendLine(("    Files           : deleted {0}, failed {1}, locked {2}, remaining {3}" -f $r.FilesDeleted, $r.FilesFailed, $r.FilesLocked, $r.FilesAfter))
        [void]$sb.AppendLine(("    Bytes reclaimed : {0}" -f (Format-ByteSize $r.BytesDeleted)))
        if ($r.BackupPath) { [void]$sb.AppendLine(("    Backup          : {0} (SHA-256 {1})" -f $r.BackupPath, $r.BackupHash)) }
        if (@($r.DependentServices).Count -gt 0) { [void]$sb.AppendLine(("    Dependents cycled: {0}" -f (@($r.DependentServices) -join ', '))) }
        foreach ($w in @($r.Warnings)) { [void]$sb.AppendLine(("    WARN : {0}" -f $w)) }
        foreach ($e in @($r.Errors)) { [void]$sb.AppendLine(("    ERROR: {0}" -f $e)) }
        if (@($r.Jobs).Count -gt 0) {
            [void]$sb.AppendLine('    Queued jobs:')
            foreach ($j in (@($r.Jobs) | Select-Object -First 50)) {
                [void]$sb.AppendLine(("      #{0,-6} {1,-28} {2,-24} {3,10}  {4}" -f $j.JobId, $j.PrinterName, $j.DocumentName, (Format-ByteSize ([long]$j.SizeBytes)), $j.Status))
            }
        }
    }
    [void]$sb.AppendLine('=' * 78)
    return $sb.ToString()
}

function ConvertTo-HtmlEncoded {
    <#
    .SYNOPSIS
        HTML-escapes any value before it is embedded in the report.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()]$Value)
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function New-HtmlReportRow {
    <#
    .SYNOPSIS
        Builds one label/value table row for the HTML report.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Key,
        [AllowNull()]$Value
    )
    return ("<tr><th>{0}</th><td>{1}</td></tr>" -f (ConvertTo-HtmlEncoded -Value $Key), (ConvertTo-HtmlEncoded -Value $Value))
}

function Get-HtmlReportStyle {
    <#
    .SYNOPSIS
        The inline stylesheet. No external resources, no script: the report must
        render identically in Outlook, a ticketing system and an air-gapped browser.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()
    return @'
:root{color-scheme:light dark}
body{font-family:"Segoe UI",Roboto,Helvetica,Arial,sans-serif;margin:0;padding:24px;background:#0f1115;color:#e6e6e6}
h1{font-size:20px;margin:0 0 4px}
h2{font-size:15px;margin:26px 0 8px;color:#9ecbff;border-bottom:1px solid #2b303b;padding-bottom:4px}
h3{font-size:13px;color:#c9d1d9;margin:14px 0 4px}
.sub{color:#9aa4b2;font-size:12px;margin-bottom:18px}
.badge{display:inline-block;padding:4px 12px;border-radius:14px;font-weight:600;font-size:13px}
.ok{background:#123d22;color:#7ee2a8}
.warn{background:#403412;color:#f2d27e}
.fail{background:#43161a;color:#ff9aa2}
.info{background:#152b40;color:#8fc7ff}
table{border-collapse:collapse;width:100%;font-size:13px;margin-bottom:6px}
th,td{text-align:left;padding:6px 10px;border-bottom:1px solid #232833;vertical-align:top}
th{color:#9aa4b2;font-weight:600;width:220px}
thead th{background:#171b23;color:#c9d1d9;width:auto}
tbody tr:hover{background:#161a22}
.num{text-align:right;font-variant-numeric:tabular-nums}
code{background:#171b23;padding:1px 5px;border-radius:4px;font-family:Consolas,monospace;font-size:12px;word-break:break-all}
.w{color:#f2d27e}
.e{color:#ff9aa2}
ul{margin:4px 0 10px 20px;padding:0}
li{margin:2px 0}
footer{margin-top:28px;color:#6b7480;font-size:11px}
'@
}

function ConvertTo-HtmlReport {
    <#
    .SYNOPSIS
        Self-contained HTML report. Every embedded value is HTML-escaped, because
        the report may end up in a ticketing system that renders it.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][AllowNull()]$Summary)

    $statusClass = switch ($Summary.OverallStatus) {
        'Success' { 'ok' }
        'NothingToDo' { 'ok' }
        'DryRun' { 'info' }
        'Diagnostic' { 'info' }
        'SuccessWithWarnings' { 'warn' }
        'PartialSuccess' { 'warn' }
        'Cancelled' { 'warn' }
        default { 'fail' }
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<!DOCTYPE html>')
    [void]$sb.AppendLine('<html lang="en"><head><meta charset="utf-8">')
    [void]$sb.AppendLine(("<title>{0} report - {1}</title>" -f (ConvertTo-HtmlEncoded -Value $Script:ScriptName), (ConvertTo-HtmlEncoded -Value $Summary.StartedUtc)))
    [void]$sb.AppendLine('<style>')
    [void]$sb.AppendLine((Get-HtmlReportStyle))
    [void]$sb.AppendLine('</style>')
    [void]$sb.AppendLine('</head><body>')
    [void]$sb.AppendLine('<h1>Print Spooler Queue Cleanup &mdash; run report</h1>')
    [void]$sb.AppendLine(('<div class="sub">Tool v{0} &middot; started {1} &middot; duration {2} &middot; operator {3} &middot; PowerShell {4}</div>' -f `
            (ConvertTo-HtmlEncoded -Value $Summary.ToolVersion), (ConvertTo-HtmlEncoded -Value $Summary.StartedUtc), `
            (ConvertTo-HtmlEncoded -Value $Summary.Duration), (ConvertTo-HtmlEncoded -Value $Summary.Operator), `
            (ConvertTo-HtmlEncoded -Value $Summary.PowerShellVersion)))
    [void]$sb.AppendLine(('<p><span class="badge {0}">{1}</span> <span class="badge info">mode: {2}</span> <span class="badge info">exit code: {3}</span></p>' -f `
            $statusClass, (ConvertTo-HtmlEncoded -Value $Summary.OverallStatus), (ConvertTo-HtmlEncoded -Value $Summary.Mode), (ConvertTo-HtmlEncoded -Value $Summary.ExitCode)))

    [void]$sb.AppendLine('<h2>Run totals</h2><table><tbody>')
    [void]$sb.AppendLine((New-HtmlReportRow -Key 'Targets attempted' -Value $Summary.TargetsAttempted))
    [void]$sb.AppendLine((New-HtmlReportRow -Key 'Succeeded' -Value $Summary.TargetsSucceeded))
    [void]$sb.AppendLine((New-HtmlReportRow -Key 'Succeeded with warnings' -Value $Summary.TargetsWithWarnings))
    [void]$sb.AppendLine((New-HtmlReportRow -Key 'Nothing to do' -Value $Summary.TargetsNothingToDo))
    [void]$sb.AppendLine((New-HtmlReportRow -Key 'Failed' -Value $Summary.TargetsFailed))
    [void]$sb.AppendLine((New-HtmlReportRow -Key 'Print jobs found / selected / removed' -Value ("{0} / {1} / {2}" -f $Summary.JobsFound, $Summary.JobsSelected, $Summary.JobsRemoved)))
    [void]$sb.AppendLine((New-HtmlReportRow -Key 'Spool files found / deleted / locked' -Value ("{0} / {1} / {2}" -f $Summary.FilesFound, $Summary.FilesDeleted, $Summary.FilesLocked)))
    [void]$sb.AppendLine((New-HtmlReportRow -Key 'Space reclaimed' -Value (Format-ByteSize $Summary.BytesDeleted)))
    [void]$sb.AppendLine((New-HtmlReportRow -Key 'Resurrected after restart' -Value $Summary.ResurrectedFiles))
    [void]$sb.AppendLine((New-HtmlReportRow -Key 'Warnings / errors' -Value ("{0} / {1}" -f $Summary.Warnings, $Summary.Errors)))
    [void]$sb.AppendLine((New-HtmlReportRow -Key 'Log file' -Value $Summary.LogPath))
    [void]$sb.AppendLine('</tbody></table>')

    [void]$sb.AppendLine('<h2>Per-computer results</h2>')
    [void]$sb.AppendLine('<table><thead><tr><th>Computer</th><th>Status</th><th class="num">Jobs</th><th class="num">Deleted</th><th class="num">Locked</th><th class="num">Reclaimed</th><th>Service now</th><th class="num">Seconds</th><th>Spool directory</th></tr></thead><tbody>')
    foreach ($r in @($Summary.Results)) {
        $cls = switch ($r.Status) {
            'Success' { 'ok' }
            'NothingToDo' { 'ok' }
            'DryRun' { 'info' }
            'Diagnostic' { 'info' }
            'SuccessWithWarnings' { 'warn' }
            default { 'fail' }
        }
        $svc = if ($r.ServiceIsRunning) { 'Running' } elseif ($null -eq $r.ServiceIsRunning) { 'unknown' } else { 'STOPPED' }
        [void]$sb.AppendLine(('<tr><td>{0}</td><td><span class="badge {1}">{2}</span></td><td class="num">{3}</td><td class="num">{4}</td><td class="num">{5}</td><td class="num">{6}</td><td>{7}</td><td class="num">{8}</td><td><code>{9}</code></td></tr>' -f `
                (ConvertTo-HtmlEncoded -Value $r.ComputerName), $cls, (ConvertTo-HtmlEncoded -Value $r.Status), `
                (ConvertTo-HtmlEncoded -Value $r.JobsRemoved), (ConvertTo-HtmlEncoded -Value $r.FilesDeleted), `
                (ConvertTo-HtmlEncoded -Value $r.FilesLocked), (ConvertTo-HtmlEncoded -Value (Format-ByteSize $r.BytesDeleted)), `
                (ConvertTo-HtmlEncoded -Value $svc), (ConvertTo-HtmlEncoded -Value ([Math]::Round($r.DurationMs / 1000, 1))), `
                (ConvertTo-HtmlEncoded -Value $r.SpoolDirectory)))
    }
    [void]$sb.AppendLine('</tbody></table>')

    $withJobs = @($Summary.Results | Where-Object { @(Get-ResultPropertyValue -Item $_ -Name 'Jobs').Count -gt 0 })
    if ($withJobs.Count -gt 0) {
        [void]$sb.AppendLine('<h2>Queued jobs discovered</h2>')
        foreach ($r in $withJobs) {
            [void]$sb.AppendLine(('<h3>{0}</h3>' -f (ConvertTo-HtmlEncoded -Value $r.ComputerName)))
            [void]$sb.AppendLine('<table><thead><tr><th>Job</th><th>Printer</th><th>Document</th><th>Owner</th><th class="num">Size</th><th class="num">Pages</th><th>Status</th><th>Submitted</th><th>Source</th></tr></thead><tbody>')
            foreach ($j in (@($r.Jobs) | Select-Object -First 200)) {
                $sub = 'n/a'
                if ($j.SubmittedUtc) {
                    try { $sub = ([DateTime]$j.SubmittedUtc).ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss') }
                    catch { $sub = [string]$j.SubmittedUtc }
                }
                [void]$sb.AppendLine(('<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td class="num">{4}</td><td class="num">{5}</td><td>{6}</td><td>{7}</td><td>{8}</td></tr>' -f `
                        (ConvertTo-HtmlEncoded -Value $j.JobId), (ConvertTo-HtmlEncoded -Value $j.PrinterName), `
                        (ConvertTo-HtmlEncoded -Value $j.DocumentName), (ConvertTo-HtmlEncoded -Value $j.UserName), `
                        (ConvertTo-HtmlEncoded -Value (Format-ByteSize ([long]$j.SizeBytes))), (ConvertTo-HtmlEncoded -Value $j.TotalPages), `
                        (ConvertTo-HtmlEncoded -Value $j.Status), (ConvertTo-HtmlEncoded -Value $sub), (ConvertTo-HtmlEncoded -Value $j.Source)))
            }
            [void]$sb.AppendLine('</tbody></table>')
        }
    }

    $withIssues = @($Summary.Results | Where-Object { @(Get-ResultPropertyValue -Item $_ -Name 'Warnings').Count -gt 0 -or @(Get-ResultPropertyValue -Item $_ -Name 'Errors').Count -gt 0 })
    if ($withIssues.Count -gt 0) {
        [void]$sb.AppendLine('<h2>Warnings and errors</h2>')
        foreach ($r in $withIssues) {
            [void]$sb.AppendLine(('<p style="margin:10px 0 2px"><strong>{0}</strong></p><ul>' -f (ConvertTo-HtmlEncoded -Value $r.ComputerName)))
            foreach ($e in @($r.Errors)) { [void]$sb.AppendLine(('<li class="e"><strong>ERROR</strong> {0}</li>' -f (ConvertTo-HtmlEncoded -Value $e))) }
            foreach ($w in @($r.Warnings)) { [void]$sb.AppendLine(('<li class="w"><strong>WARN</strong> {0}</li>' -f (ConvertTo-HtmlEncoded -Value $w))) }
            [void]$sb.AppendLine('</ul>')
        }
    }

    $backups = @($Summary.Results | Where-Object { -not [string]::IsNullOrWhiteSpace([string](Get-ResultPropertyValue -Item $_ -Name 'BackupPath')) })
    if ($backups.Count -gt 0) {
        [void]$sb.AppendLine('<h2>Backups created</h2><table><thead><tr><th>Computer</th><th>Archive</th><th>SHA-256</th></tr></thead><tbody>')
        foreach ($b in $backups) {
            [void]$sb.AppendLine(('<tr><td>{0}</td><td><code>{1}</code></td><td><code>{2}</code></td></tr>' -f `
                    (ConvertTo-HtmlEncoded -Value $b.ComputerName), (ConvertTo-HtmlEncoded -Value $b.BackupPath), (ConvertTo-HtmlEncoded -Value $b.BackupHash)))
        }
        [void]$sb.AppendLine('</tbody></table>')
        [void]$sb.AppendLine('<p class="sub">Spool files cannot be re-queued by copying them back into PRINTERS. Use a backup to recover the document payload and re-submit it from the application.</p>')
    }

    [void]$sb.AppendLine(('<footer>Generated by {0} v{1} &middot; {2} &middot; {3}</footer>' -f `
            (ConvertTo-HtmlEncoded -Value $Script:ScriptName), (ConvertTo-HtmlEncoded -Value $Script:ScriptVersion), `
            (ConvertTo-HtmlEncoded -Value $Script:ScriptRepoUrl), (ConvertTo-HtmlEncoded -Value (Get-Date).ToString('u'))))
    [void]$sb.AppendLine('</body></html>')
    return $sb.ToString()
}

function Export-RunReport {
    <#
    .SYNOPSIS
        Writes the run report to disk in one or more formats.
    .DESCRIPTION
        -Path may be a directory (file names are generated) or a full file path
        (used for the first requested format). Returns the list of files written.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory)][AllowNull()]$Summary,
        # Empty ELEMENTS are tolerated on purpose. A format list can arrive with a
        # blank entry (an unset or empty -ExportFormat value, or a stray separator
        # in the configured list). Without AllowEmptyString() the binder rejects the
        # whole call with "Cannot bind argument to parameter 'Format' because it is
        # an empty string", which aborted the run instead of simply writing nothing
        # for that entry. Blanks are dropped below.
        [Parameter(Mandatory)][AllowNull()][AllowEmptyCollection()][AllowEmptyString()][string[]]$Format,
        [AllowNull()][AllowEmptyString()][string]$Path
    )

    $written = New-Object System.Collections.Generic.List[string]
    $Format = @($Format | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if (-not $Format -or @($Format).Count -eq 0) { return $written.ToArray() }

    $stamp = Get-SafeFileNameStamp
    $isDir = $false
    if ([string]::IsNullOrWhiteSpace($Path)) {
        $Path = $Script:RuntimeState.ScriptDirectory
        $isDir = $true
    }
    elseif (Test-Path -LiteralPath $Path -PathType Container) { $isDir = $true }
    elseif ($Path.EndsWith('\') -or $Path.EndsWith('/')) {
        $isDir = $true
        if (-not (Test-Path -LiteralPath $Path)) { New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop | Out-Null }
    }

    $Format = @($Format | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    foreach ($f in @($Format)) {
        $fmt = switch (([string]$f).ToLowerInvariant()) {
            'json' { 'Json' }
            'csv' { 'Csv' }
            'html' { 'Html' }
            'htm' { 'Html' }
            default { 'Text' }
        }
        $ext = switch ($fmt) { 'Json' { 'json' } 'Csv' { 'csv' } 'Html' { 'html' } default { 'txt' } }

        if ($isDir) {
            $target = Join-Path $Path ("{0}_{1}.{2}" -f $Script:ScriptName, $stamp, $ext)
        }
        elseif ($written.Count -eq 0) {
            $target = $Path
        }
        else {
            $target = Join-Path (Split-Path -Parent $Path) ("{0}_{1}.{2}" -f [IO.Path]::GetFileNameWithoutExtension($Path), $fmt.ToLowerInvariant(), $ext)
        }

        $dir = Split-Path -Parent $target
        if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null }

        try {
            $content = switch ($fmt) {
                'Json' { ConvertTo-JsonReport -Summary $Summary }
                'Csv' { ConvertTo-CsvReport -Summary $Summary }
                'Html' { ConvertTo-HtmlReport -Summary $Summary }
                default { ConvertTo-TextReport -Summary $Summary }
            }
            Set-Content -LiteralPath $target -Value $content -Encoding UTF8 -ErrorAction Stop
            $written.Add($target)
            Write-LogSuccess "Report written: $target" -Component 'Report'
        }
        catch {
            Write-LogError "Report export ($fmt) failed for '$target': $($_.Exception.Message)" -Component 'Report'
        }
    }
    return $written.ToArray()
}

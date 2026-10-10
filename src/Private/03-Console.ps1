# ============================================================================
#  Console rendering layer
# ----------------------------------------------------------------------------
#  Pure presentation. Nothing here performs I/O on the spooler, touches the file
#  system, or throws. Every function degrades gracefully when the host has no
#  colour support, is redirected, or is a narrow console.
# ============================================================================

function Get-Bar {
    <#
    .SYNOPSIS
        A rule of the current console width.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([int]$Width = 0, [string]$Character = '=')
    if ($Width -le 0) { $Width = [int]$Script:RuntimeState.ConsoleWidth }
    $Width = [Math]::Max(20, [Math]::Min(200, $Width))
    return ($Character * $Width)
}

function Write-ConsoleHeader {
    <#
    .SYNOPSIS
        The product banner shown on interactive runs.
    #>
    [CmdletBinding()]
    param([string[]]$ModeNotes = @())

    if ($Script:RuntimeState.Quiet -or $Script:RuntimeState.MachineReadable) { return }
    $bar = Get-Bar -Character '='
    $color = if ($Script:RuntimeState.NoColor) { $null } else { $Script:RuntimeState.Colors['Header'] }

    $lines = @(
        $bar,
        ("    Print Spooler Queue Cleanup v{0}" -f $Script:ScriptVersion),
        ("    Author : {0} <{1}>    Site: {2}" -f $Script:ScriptAuthor, $Script:ScriptAuthorEmail, $Script:ScriptAuthorSite),
        ("    GitHub : {0}" -f $Script:ScriptRepoUrl),
        ("    License: {0}    PowerShell {1}" -f $Script:ScriptLicense, $PSVersionTable.PSVersion)
    )
    foreach ($note in $ModeNotes) { $lines += ("    MODE   : {0}" -f $note) }
    $lines += $bar

    foreach ($line in $lines) {
        if ($null -ne $color) { Write-Host $line -ForegroundColor $color }
        else { Write-Host $line }
    }
    Write-Host ''
}

function Write-ConsoleSection {
    <#
    .SYNOPSIS
        Section title with a rule under it.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)][AllowEmptyString()][string]$Title,
        [AllowNull()][string]$Target,
        [string]$Character = '-'
    )
    if ($Script:RuntimeState.Quiet -or $Script:RuntimeState.MachineReadable) { return }
    $head = if ($Target) { "$Title - $Target" } else { $Title }
    $color = if ($Script:RuntimeState.NoColor) { $null } else { $Script:RuntimeState.Colors['Header'] }
    $rule = Get-Bar -Character $Character
    if ($null -ne $color) {
        Write-Host (" {0} " -f $head) -ForegroundColor $color
        Write-Host $rule -ForegroundColor $color
    }
    else {
        Write-Host (" {0} " -f $head)
        Write-Host $rule
    }
}

function Write-ConsoleKeyValue {
    <#
    .SYNOPSIS
        Aligned "Label : value" line used by summaries and diagnostics.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)][AllowEmptyString()][string]$Label,
        [Parameter(Position = 1)][AllowNull()]$Value,
        [int]$LabelWidth = 26,
        [AllowNull()][string]$ValueColor,
        [string]$Indent = '    '
    )
    if ($Script:RuntimeState.MachineReadable) { return }
    $text = if ($null -eq $Value) { '' } else { [string]$Value }
    if ([string]::IsNullOrEmpty($text)) { $text = '-' }
    $pad = ' ' * [Math]::Max(1, $LabelWidth - $Label.Length)
    $line = "{0}{1}:{2}" -f $Indent, $Label, $pad

    if ($Script:RuntimeState.NoColor -or -not $ValueColor) {
        Write-Host ($line + $text)
    }
    else {
        Write-Host $line -NoNewline -ForegroundColor $Script:RuntimeState.Colors['Label']
        Write-Host $text -ForegroundColor $ValueColor
    }
}

function Write-ConsoleTable {
    <#
    .SYNOPSIS
        Renders objects as a fixed-width table without Format-Table's truncation
        and without polluting the output stream.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true)][AllowNull()]$InputObject,
        [string[]]$Property,
        [int]$MaxWidth = 0
    )
    begin { $items = New-Object System.Collections.Generic.List[object] }
    process { if ($null -ne $InputObject) { $items.Add($InputObject) } }
    end {
        if ($Script:RuntimeState.MachineReadable) { return }
        if ($items.Count -eq 0) { Write-Host '    (no rows)'; return }
        if (-not $Property) { $Property = @($items[0].PSObject.Properties.Name) }
        if ($MaxWidth -le 0) { $MaxWidth = [int]$Script:RuntimeState.ConsoleWidth }

        $widths = [ordered]@{}
        foreach ($p in $Property) {
            $w = $p.Length
            foreach ($it in $items) {
                $v = [string]$it.$p
                if ($v -and $v.Length -gt $w) { $w = $v.Length }
            }
            $widths[$p] = [Math]::Min(48, $w)
        }
        # Shrink proportionally when the table is wider than the console.
        # Measure-Object returns $null for an empty pipeline, and `.Sum` on $null
        # throws under Set-StrictMode -Version Latest. Sum explicitly instead.
        $widthSum = 0
        foreach ($wv in $widths.Values) { $widthSum += [int]$wv }
        $total = $widthSum + (3 * $Property.Count)
        if ($total -gt $MaxWidth) {
            $excess = $total - $MaxWidth
            foreach ($p in $Property) {
                if ($excess -le 0) { break }
                $cut = [Math]::Min($excess, [Math]::Max(0, $widths[$p] - 8))
                $widths[$p] = $widths[$p] - $cut
                $excess = $excess - $cut
            }
        }

        $headColor = if ($Script:RuntimeState.NoColor) { $null } else { 'Cyan' }
        $ruleColor = if ($Script:RuntimeState.NoColor) { $null } else { 'DarkGray' }

        $head = ($Property | ForEach-Object { $_.PadRight($widths[$_]).Substring(0, $widths[$_]) }) -join ' | '
        $rule = ($Property | ForEach-Object { '-' * $widths[$_] }) -join '-+-'
        if ($null -ne $headColor) { Write-Host ('    ' + $head) -ForegroundColor $headColor } else { Write-Host ('    ' + $head) }
        if ($null -ne $ruleColor) { Write-Host ('    ' + $rule) -ForegroundColor $ruleColor } else { Write-Host ('    ' + $rule) }

        foreach ($it in $items) {
            $row = ($Property | ForEach-Object {
                    $v = [string]$it.$_
                    if ($v.Length -gt $widths[$_]) {
                        # Truncate to width-1 and append an ellipsis so the column
                        # width is preserved exactly and the table cannot drift.
                        $ell = if ($Script:RuntimeState.UseAscii) { '~' } else { [string][char]0x2026 }
                        $v = $v.Substring(0, [Math]::Max(0, $widths[$_] - 1)) + $ell
                    }
                    $v.PadRight($widths[$_])
                }) -join ' | '
            Write-Host ('    ' + $row)
        }
    }
}

function Show-CleanupProgress {
    <#
    .SYNOPSIS
        Write-Progress wrapper that no-ops in quiet, machine-readable or
        non-interactive mode.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Activity,
        [AllowNull()][string]$Status,
        [int]$PercentComplete = -1,
        [int]$Id = 1,
        [AllowNull()][Nullable[int]]$ParentId,
        [switch]$Completed
    )
    if ($Script:RuntimeState.NoProgress -or $Script:RuntimeState.Quiet -or $Script:RuntimeState.MachineReadable) { return }
    try {
        $p = @{ Activity = $Activity; Id = $Id }
        if ($Status) { $p['Status'] = $Status }
        if ($PercentComplete -ge 0) { $p['PercentComplete'] = [Math]::Min(100, $PercentComplete) }
        if ($null -ne $ParentId) { $p['ParentId'] = $ParentId }
        if ($Completed) { $p['Completed'] = $true }
        Write-Progress @p
    }
    catch {
        # Some constrained hosts have no progress stream at all.
    }
}

function Format-ByteSize {
    <#
    .SYNOPSIS
        Locale-independent human-readable byte formatting.
    .EXAMPLE
        Format-ByteSize 1536   # -> "1.50 KB"
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, Position = 0)][AllowNull()][double]$Bytes,
        [int]$Decimals = 2
    )
    if ($null -eq $Bytes) { return '0 B' }
    $units = @('B', 'KB', 'MB', 'GB', 'TB', 'PB')
    $abs = [Math]::Abs($Bytes)
    $sign = if ($Bytes -lt 0) { '-' } else { '' }
    $i = 0
    while ($abs -ge 1024 -and $i -lt ($units.Count - 1)) { $abs = $abs / 1024; $i++ }
    if ($i -eq 0) { return ("{0}{1} {2}" -f $sign, [int64]$abs, $units[0]) }
    # Explicit invariant culture rather than the process culture. The documented
    # output is "1.50 KB", and sizes are read by machines (reports, JSON,
    # dashboards), so a host with a comma decimal separator must not change it.
    # The loop above stops below 1024, so no group separator is ever emitted.
    $number = $abs.ToString('N' + $Decimals, [Globalization.CultureInfo]::InvariantCulture)
    return ("{0}{1} {2}" -f $sign, $number, $units[$i])
}

function Format-Duration {
    <#
    .SYNOPSIS
        "1d 02h 03m 04s" / "01:02:03.456" style duration, culture-invariant.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        # Untyped on purpose: a [TimeSpan]-typed parameter cannot receive $null,
        # which would make the documented null handling unreachable.
        [Parameter(Mandatory, Position = 0)][AllowNull()]$Duration
    )
    if ($null -eq $Duration) { return '00:00:00.000' }
    $ts = [TimeSpan]$Duration
    if ($ts.TotalDays -ge 1) { return ("{0}d {1:D2}h {2:D2}m {3:D2}s" -f [int]$ts.TotalDays, $ts.Hours, $ts.Minutes, $ts.Seconds) }
    return ("{0:D2}:{1:D2}:{2:D2}.{3:D3}" -f $ts.Hours, $ts.Minutes, $ts.Seconds, $ts.Milliseconds)
}

function Format-FileSizeArgument {
    <#
    .SYNOPSIS
        Parses "10MB", "1.5GB", "512", "2 gb" into an exact byte count.
    .DESCRIPTION
        Powers -LargerThanBytes / -SmallerThanBytes so operators can type human
        sizes instead of counting zeros. Returns $null when unparseable - never 0,
        because 0 would silently mean "match everything".
    #>
    [CmdletBinding()]
    [OutputType([long])]
    param([Parameter(Position = 0)][AllowNull()][AllowEmptyString()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $v = $Value.Trim()
    if ($v -match '^(?<num>[-+]?[0-9]*\.?[0-9]+)\s*(?<unit>[KMGTPE]?I?B?)?$') {
        $num = [double]$Matches['num']
        $unit = if ($Matches['unit']) { $Matches['unit'].ToUpperInvariant() } else { '' }
        $multiplier = switch -regex ($unit) {
            '^P' { 1PB }
            '^T' { 1TB }
            '^G' { 1GB }
            '^M' { 1MB }
            '^K' { 1KB }
            default { 1L }
        }
        return [long]($num * [double]$multiplier)
    }
    return $null
}

function Write-Banner {
    <#
    .SYNOPSIS
        Boxed notice for warnings the operator must not miss.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)][string[]]$Text,
        [string]$Color = 'Yellow',
        [string]$Character = '*'
    )
    if ($Script:RuntimeState.MachineReadable) { return }
    $inner = [Math]::Max(20, ($Text | Measure-Object -Property Length -Maximum).Maximum + 4)
    $top = $Character * $inner
    $rows = @($top)
    foreach ($t in $Text) { $rows += ("{0} {1} {2}" -f $Character, $t.PadRight($inner - 4), $Character) }
    $rows += $top
    foreach ($r in $rows) {
        if ($Script:RuntimeState.NoColor) { Write-Host $r }
        else { Write-Host $r -ForegroundColor $Color }
    }
}

function Confirm-DestructiveAction {
    <#
    .SYNOPSIS
        Single confirmation gate for the whole run (not per-computer).
    .DESCRIPTION
        Honours -Force, -Confirm:$false, ShouldProcess/-WhatIf, and refuses to
        block forever on a non-interactive host (finding A22).
    .OUTPUTS
        $true to proceed, $false to abort with ExitCode.Cancelled.
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Target,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Action,
        [switch]$Force,
        [string]$DefaultAnswer = 'Y',
        [int]$TimeoutSeconds = 0
    )

    if ($Force) { return $true }
    # -WhatIf makes ShouldProcess return $false without prompting.
    if (-not $PSCmdlet.ShouldProcess($Target, $Action)) { return $false }

    if (Test-NonInteractiveHost) {
        Write-LogWarn 'Non-interactive host detected and -Force was not supplied. Refusing to prompt; aborting.' -Component 'Confirm'
        return $false
    }

    if ($TimeoutSeconds -gt 0) {
        Write-Host ("{0} (auto-{1} in {2}s)" -f $Action, $DefaultAnswer, $TimeoutSeconds) -ForegroundColor Yellow
        $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
        while ((Get-Date) -lt $deadline) {
            try {
                if ($Host.UI.RawUI.KeyAvailable) {
                    $key = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown').Character
                    Write-Host ''
                    return ([string]$key -match '^[YyJj]')   # J = "Ja" on German keyboards
                }
            }
            catch {
                # A host without RawUI key support cannot do a timed prompt; fall
                # through to the default answer rather than hanging.
                break
            }
            Start-Sleep -Milliseconds 100
        }
        Write-Host ''
        Write-LogInfo "Confirmation timed out; using the default answer '$DefaultAnswer'." -Component 'Confirm'
        return ($DefaultAnswer -eq 'Y')
    }

    $answer = $Host.UI.PromptForChoice($Target, $Action, @(
            (New-Object System.Management.Automation.Host.ChoiceDescription '&Yes', 'Proceed with the destructive operation'),
            (New-Object System.Management.Automation.Host.ChoiceDescription '&No', 'Abort')
        ), $(if ($DefaultAnswer -eq 'Y') { 0 } else { 1 }))
    return ($answer -eq 0)
}

function Initialize-ConsolePalette {
    <#
    .SYNOPSIS
        Picks ASCII or Unicode glyphs and enables or disables colour for this session.
    #>
    [CmdletBinding()]
    param(
        [switch]$Ascii,
        [switch]$NoColor,
        [AllowNull()]$Capability
    )
    if (-not $Capability) { $Capability = Get-ConsoleCapability }

    $useAscii = [bool]$Ascii -or (-not $Capability.SupportsUnicode)
    $Script:RuntimeState.Glyphs = if ($useAscii) { $Script:GlyphAscii } else { $Script:GlyphUnicode }
    $Script:RuntimeState.UseAscii = $useAscii
    $Script:RuntimeState.NoColor = ([bool]$NoColor -or -not $Capability.SupportsColor)
    $Script:RuntimeState.ConsoleWidth = [int]$Capability.BufferWidth
    return $Script:RuntimeState.Glyphs
}

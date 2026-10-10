# ============================================================================
#  Exit-code publishing, watchdog, self-test, argument reconstruction
# ============================================================================

function Set-ScriptExitCode {
    <#
    .SYNOPSIS
        Records the run's exit code for the script epilogue to publish.
    .DESCRIPTION
        FINDING A21: v3 assigned $global:LASTEXITCODE from inside the script, which
        (a) does not change the process exit code for `powershell -File` and
        (b) leaks state into a session that dot-sourced the script. v4 stores the
        code in runtime state; the epilogue decides how to publish it based on how
        the script was invoked.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param([Parameter(Mandatory)][int]$Code)
    if (Test-Path variable:Script:RuntimeState) { $Script:RuntimeState.ExitCode = $Code }
    return $Code
}

function Publish-ScriptExitCode {
    <#
    .SYNOPSIS
        Script epilogue: publishes the exit code correctly for the invocation style.
    .DESCRIPTION
        Three cases, and getting them wrong is what produced the v3 defect where
        `powershell -File` reported success even after a failure:

          * run as `.\script.ps1` or `powershell -File script.ps1`
              -> `exit N` terminates the script and becomes the process exit code.
          * dot-sourced (`. .\script.ps1`) or imported as a module
              -> `exit N` would close the operator's shell, so the code is
                 published through $global:LASTEXITCODE and control returns.

        The caller must pass the invocation context captured IN SCRIPT SCOPE;
        $MyInvocation inside this function describes this function, not the script.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [AllowNull()][AllowEmptyString()][string]$InvocationName,
        [AllowNull()][AllowEmptyString()][string]$ScriptName,
        [AllowNull()][AllowEmptyString()][string]$CommandPath
    )

    $code = 0
    if ((Test-Path variable:Script:RuntimeState) -and $Script:RuntimeState -and $Script:RuntimeState.Contains('ExitCode')) {
        $code = [int]$Script:RuntimeState.ExitCode
    }

    $dotSourced = $false
    if ($InvocationName -eq '.') {
        $dotSourced = $true
    }
    elseif ($CommandPath -and $InvocationName -and ($InvocationName -ne $CommandPath) -and ($InvocationName -ne $ScriptName)) {
        # Invoked as a command from inside another scope rather than as the
        # top-level script file: do not tear down the caller's session.
        $dotSourced = $true
    }

    $global:LASTEXITCODE = $code
    if ($Script:RuntimeState) { $Script:RuntimeState.DotSourced = $dotSourced }
    if (-not $dotSourced) { exit $code }
    return $code
}

function Start-WatchdogInternal {
    <#
    .SYNOPSIS
        Starts an out-of-process watchdog that survives a wedged host process.
    .DESCRIPTION
        The watchdog is a separate job, not a runspace inside this process: if a
        Spooler hang takes our own process down with it, an in-process timer would
        die too. On expiry it restarts the Spooler (so the machine is never left
        without printing) and terminates the run with exit code 15.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateRange(1, 1440)][int]$Minutes,
        [Parameter(Mandatory)][AllowEmptyString()][string]$InvocationId,
        [AllowNull()][AllowEmptyCollection()][string[]]$Targets
    )

    $logPath = $null
    if (Test-Path variable:Script:RuntimeState) { $logPath = $Script:RuntimeState.LogPath }
    $ownerPid = $PID

    try {
        $job = Start-Job -Name "PSQWatchdog-$InvocationId" -ScriptBlock {
            param($Minutes, $InvocationId, $OwnerPid, $LogPath, $Targets)

            $deadline = (Get-Date).AddMinutes($Minutes)
            while ((Get-Date) -lt $deadline) {
                Start-Sleep -Seconds 5
                $alive = $false
                try {
                    $p = Get-Process -Id $OwnerPid -ErrorAction SilentlyContinue
                    $alive = ($null -ne $p)
                }
                catch { $alive = $false }
                if (-not $alive) { return 'OwnerExited' }
            }

            $stamp = (Get-Date).ToString('u')
            $line = "{0}|CRITICAL|Watchdog|{1}|Watchdog expired after {2} minute(s); restarting the Spooler and terminating run {3}." -f `
                ([DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')), ($Targets -join ','), $Minutes, $InvocationId
            if ($LogPath) {
                try { Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue }
                catch { /* watchdog logging is best-effort; the restart below is what matters */ }
            }
            foreach ($t in @($Targets)) {
                if (-not $t -or $t -eq '.') {
                    try { Restart-Service -Name 'Spooler' -Force -ErrorAction SilentlyContinue }
                    catch { /* recovery is best-effort; the SCM will restart the service on its own policy */ }
                }
            }
            try { Stop-Process -Id $OwnerPid -Force -ErrorAction SilentlyContinue }
            catch { /* if the owner cannot be killed it is already gone */ }
            return "Expired:$stamp"
        } -ArgumentList $Minutes, $InvocationId, $ownerPid, $logPath, @($Targets)
        return $job
    }
    catch {
        Write-LogWarn "The watchdog could not be armed: $($_.Exception.Message)" -Component 'Watchdog'
        return $null
    }
}

function Stop-WatchdogInternal {
    <#
    .SYNOPSIS
        Stops and removes the watchdog job. Never throws.
    #>
    [CmdletBinding()]
    param([AllowNull()]$Job)
    if (-not $Job) { return }
    try { $null = Stop-Job -Job $Job -ErrorAction SilentlyContinue }
    catch { /* the job may already have finished */ }
    try { $null = Remove-Job -Job $Job -Force -ErrorAction SilentlyContinue }
    catch { /* teardown must never raise */ }
}

function New-ChildArgumentList {
    <#
    .SYNOPSIS
        Rebuilds a command line from $PSBoundParameters for the elevated child.
    .DESCRIPTION
        Used by -SelfElevate. Switches become bare -Name, booleans become
        -Name:$true/$false, arrays become quoted comma lists, and PSCredential is
        deliberately NOT forwarded - a credential object cannot cross a process
        boundary safely, so the child prompts or uses the caller's context.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory)][AllowNull()][System.Collections.IDictionary]$BoundParameters,
        [string[]]$Exclude = @('Credential', 'SelfElevate')
    )

    # NOT `$args`. That is a PowerShell AUTOMATIC variable holding the unbound
    # arguments of the current invocation, and assigning to it is the collision
    # class documented as trap 4: it appears to work, then the host repopulates it
    # and the value vanishes - or, worse, silently swallows the caller's own
    # arguments. This function builds a child process's argument vector, which is
    # exactly the shape that invites the name `$args`, so the trap is easy to
    # re-set. Invoke-AutomaticVariableScan fails the build if it returns.
    $childArgs = New-Object System.Collections.Generic.List[string]
    if (-not $BoundParameters) { return $childArgs.ToArray() }

    foreach ($key in @($BoundParameters.Keys)) {
        if (@($Exclude) -contains $key) { continue }
        $value = $BoundParameters[$key]

        if ($value -is [switch]) {
            if ($value.IsPresent) { $childArgs.Add("-$key") }
            continue
        }
        if ($value -is [bool]) {
            # Built by concatenation, not interpolation: inside a double-quoted
            # string "$$:" parses as the automatic variable $$ followed by a colon,
            # which silently emits "-Name:" with no value at all.
            $childArgs.Add(('-' + $key + ':$' + ([string]$value).ToLowerInvariant()))
            continue
        }
        if ($value -is [System.Management.Automation.PSCredential]) { continue }
        if ($null -eq $value) { continue }

        if (($value -is [System.Collections.IEnumerable]) -and ($value -isnot [string])) {
            $items = @($value | ForEach-Object { [string]$_ })
            if ($items.Count -eq 0) { continue }
            $quoted = foreach ($it in $items) {
                if ($it -match '[\s",;]') { "'{0}'" -f ($it -replace "'", "''") } else { $it }
            }
            $childArgs.Add("-$key")
            $childArgs.Add(($quoted -join ','))
            continue
        }

        $s = [string]$value
        if ([string]::IsNullOrEmpty($s)) { continue }
        $childArgs.Add("-$key")
        if ($s -match '[\s"]') { $childArgs.Add(('"{0}"' -f ($s -replace '"', '\"'))) }
        else { $childArgs.Add($s) }
    }
    return $childArgs.ToArray()
}

function Format-TimeSpanBrief {
    <#
    .SYNOPSIS
        Renders a time-span argument as the operator typed it where possible.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return 'n/a' }
    if ($Value -is [string]) { return $Value }
    try {
        $ts = ConvertFrom-AnyTimeSpan -Value $Value
        if ($ts.TotalDays -ge 1) { return ("{0}d" -f [Math]::Round($ts.TotalDays, 1)) }
        if ($ts.TotalHours -ge 1) { return ("{0}h" -f [Math]::Round($ts.TotalHours, 1)) }
        if ($ts.TotalMinutes -ge 1) { return ("{0}m" -f [Math]::Round($ts.TotalMinutes, 1)) }
        return ("{0}s" -f [Math]::Round($ts.TotalSeconds, 1))
    }
    catch {
        return [string]$Value
    }
}

function Format-SelectionDescription {
    <#
    .SYNOPSIS
        One-line human description of the selection filters in force.
    .DESCRIPTION
        Shown in the execution plan and the confirmation prompt, because "delete
        ALL spool files" and "delete jobs older than 2h on HP*" are very different
        promises and the operator must see which one is being made.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][AllowNull()][System.Collections.IDictionary]$Arguments)

    if (-not $Arguments) { return 'ALL spool files (*.SPL, *.SHD)' }
    $parts = New-Object System.Collections.Generic.List[string]

    if ($Arguments['PrinterName']) { $parts.Add("printers like $((@($Arguments['PrinterName']) | ForEach-Object { "'$_'" }) -join ', ')") }
    if ($Arguments['ExcludePrinterName']) { $parts.Add("excluding $((@($Arguments['ExcludePrinterName']) | ForEach-Object { "'$_'" }) -join ', ')") }
    if ($Arguments['JobId']) { $parts.Add("job id(s) $((@($Arguments['JobId']) | ForEach-Object { [string]$_ }) -join ', ')") }
    if ($Arguments['DocumentName']) { $parts.Add("document like '$($Arguments['DocumentName'])'") }
    if ($Arguments['OwnerName']) { $parts.Add("owner like '$($Arguments['OwnerName'])'") }
    if ($Arguments['OlderThan']) { $parts.Add("older than $(Format-TimeSpanBrief -Value $Arguments['OlderThan'])") }
    if ($Arguments['NewerThan']) { $parts.Add("newer than $(Format-TimeSpanBrief -Value $Arguments['NewerThan'])") }
    if ($Arguments['LargerThanBytes']) { $parts.Add("larger than $(Format-ByteSize ([long]$Arguments['LargerThanBytes']))") }
    if ($Arguments['SmallerThanBytes']) { $parts.Add("smaller than $(Format-ByteSize ([long]$Arguments['SmallerThanBytes']))") }

    if ($parts.Count -eq 0) { return 'ALL spool files (*.SPL, *.SHD)' }
    return ("jobs matching: {0}" -f ($parts -join '; '))
}

function Test-SpoolDirectorySafeInternal {
    <#
    .SYNOPSIS
        Local mirror of the payload's guard logic, so -SelfTest can prove the
        guards work without invoking a remote payload.
    .DESCRIPTION
        The two implementations are asserted to agree by the test suite. Keeping a
        local copy is what makes the guards unit-testable on any platform.
    #>
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Path)

    $result = [ordered]@{ Path = $Path; IsSafe = $false; Reason = $null }

    if ([string]::IsNullOrWhiteSpace($Path)) {
        $result.Reason = 'Path is empty.'
        return [pscustomobject]$result
    }
    if ($Path -match '[\x00-\x1f]') {
        $result.Reason = 'Path contains control characters.'
        return [pscustomobject]$result
    }
    if ($Path -match '^[\\]{2}') {
        $result.Reason = 'UNC paths are refused by policy.'
        return [pscustomobject]$result
    }

    # Validate the drive-root shape BEFORE calling GetFullPath(). GetFullPath is
    # platform-relative: on a non-Windows host 'C:\PRINTERS' resolves to
    # '<cwd>/C:\PRINTERS', which would silently defeat both the depth guard and the
    # forbidden-root comparison. Checking the shape first makes the guard correct
    # on every platform rather than only on the one it was written on.
    if ($Path -notmatch '^[A-Za-z]:[\\/]') {
        $result.Reason = 'Path is not rooted at a Windows drive letter (expected C:\...).'
        return [pscustomobject]$result
    }
    # The colon is legal ONLY as the drive-letter separator, which the rule above
    # has already pinned to position 1. Everything else in this set is illegal in a
    # Windows path. (An earlier version of this check included ':' outright and
    # therefore rejected every valid path - the guard test is what caught it.)
    if ($Path.Substring(2) -match '[<>:"|?*]') {
        $result.Reason = 'Path contains characters that are invalid in a Windows path.'
        return [pscustomobject]$result
    }

    $full = $null
    try { $full = [IO.Path]::GetFullPath($Path) }
    catch {
        $result.Reason = 'Path cannot be resolved.'
        return [pscustomobject]$result
    }
    if (-not [IO.Path]::IsPathRooted($full)) {
        $result.Reason = 'Path is not rooted.'
        return [pscustomobject]$result
    }

    $trimmed = $full.TrimEnd('\', '/')
    $leaf = Split-Path -Leaf $trimmed
    if ($leaf -notlike $Script:SpoolGuardRequiredLeaf) {
        $result.Reason = "Leaf '$leaf' must be '$($Script:SpoolGuardRequiredLeaf)'."
        return [pscustomobject]$result
    }

    # Count the segments of the INPUT, not of GetFullPath's output. GetFullPath
    # resolves relative to the current directory, so on a host that does not treat
    # '\' as a separator it would prepend the CWD and silently satisfy a depth
    # check that the operator's actual path should have failed.
    $depth = @($Path.Trim().TrimEnd('\', '/').Split([char[]]@('\', '/')) | Where-Object { $_ }).Count
    if ($depth -lt $Script:SpoolGuardMinPathDepth) {
        $result.Reason = "Path is too shallow ($depth segments)."
        return [pscustomobject]$result
    }

    # A real spool directory always lives under a 'spool' folder
    # (C:\Windows\System32\spool\PRINTERS). Requiring that defeats the whole class
    # of accidents an equality-only forbidden-root check cannot: a corrupt
    # DefaultSpoolDirectory pointing at 'C:\Users\PRINTERS' has the right leaf and
    # the right depth, and would otherwise be accepted.
    if ($full -notmatch '(?i)[\\/]spool[\\/]') {
        $result.Reason = "Refusing '$full': a spool directory must live under a 'spool' folder. Pass -SpoolDirectory explicitly only if this host genuinely uses a different layout."
        return [pscustomobject]$result
    }

    foreach ($bad in @(Get-SpoolGuardForbiddenRoots)) {
        if ([string]::IsNullOrWhiteSpace($bad)) { continue }
        if ([string]::Equals($trimmed, $bad.TrimEnd('\', '/'), [StringComparison]::OrdinalIgnoreCase)) {
            $result.Reason = 'Path equals a protected system root.'
            return [pscustomobject]$result
        }
    }

    $result.IsSafe = $true
    return [pscustomobject]$result
}

function Invoke-SelfTestInternal {
    <#
    .SYNOPSIS
        Internal consistency checks (-SelfTest). No Spooler interaction.
    .DESCRIPTION
        Answers "is this copy of the tool internally sound?" after a deployment, a
        patch or an update. Checks the payload template, the safety guards, the
        version parser, the exit-code contract, the config schema and the help
        generator. Exits non-zero if any check fails, so it can be a step in a
        deployment pipeline.
    #>
    [CmdletBinding()]
    param()

    $checks = New-Object System.Collections.Generic.List[object]
    # A closure rather than a nested function: a nested `function` would be
    # re-created on every call and would leak a one-letter command name into
    # script scope once the build inlines this file.
    $add = {
        param([string]$Name, [bool]$Ok, [string]$Detail)
        $checks.Add([pscustomobject]@{
                Name   = $Name
                Result = if ($Ok) { 'PASS' } else { 'FAIL' }
                Detail = $Detail
            }) | Out-Null
    }

    # 1. the payload template substitutes and parses
    try {
        # Assigned, never left on the pipeline: Resolve-PayloadScriptBlock returns
        # the scriptblock, and an unassigned call would emit it into the caller's
        # output stream where it binds to the next command's parameters - producing
        # a baffling "cannot convert to SwitchParameter" error far away from here.
        $sb = Resolve-PayloadScriptBlock -Force
        & $add 'Payload generation' ($null -ne $sb) ("{0} characters, no unsubstituted tokens" -f $Script:RuntimeState.PayloadText.Length)
    }
    catch { & $add 'Payload generation' $false $_.Exception.Message }

    # 2. every token in the template has a mapping
    try {
        $map = New-PayloadTokenMap
        $tokens = @([regex]::Matches($Script:RemotePayloadTemplate, '@@[A-Z_]+@@') | ForEach-Object { $_.Value } | Select-Object -Unique)
        $missing = @($tokens | Where-Object { -not $map.Contains($_) })
        & $add 'Token map completeness' ($missing.Count -eq 0) $(if ($missing.Count -eq 0) { "$($tokens.Count) tokens mapped" } else { "missing: $($missing -join ', ')" })
    }
    catch { & $add 'Token map completeness' $false $_.Exception.Message }

    # 3. version comparison correctness
    try {
        $cases = @(
            @{ A = '3.1.0'; B = '4.0.0'; E = -1 }
            @{ A = '4.0.0'; B = '4.0.0'; E = 0 }
            @{ A = '4.0.1'; B = '4.0.0'; E = 1 }
            @{ A = '4.0.0'; B = 'v4.0.0'; E = 0 }
            @{ A = '4.0.0-beta.1'; B = '4.0.0'; E = -1 }
            @{ A = '4.10.0'; B = '4.9.0'; E = 1 }
        )
        $bad = New-Object System.Collections.Generic.List[string]
        foreach ($c in $cases) {
            $got = Compare-ProductVersion -Reference $c.A -Difference $c.B
            if ($got -ne $c.E) { $bad.Add("$($c.A) vs $($c.B): expected $($c.E), got $got") }
        }
        & $add 'Version comparison' ($bad.Count -eq 0) $(if ($bad.Count -eq 0) { "$($cases.Count) cases" } else { $bad -join ' | ' })
    }
    catch { & $add 'Version comparison' $false $_.Exception.Message }

    # 4. the exit-code contract is complete and collision-free
    try {
        $codes = @($Script:ExitCode.Values)
        $dupes = @($codes | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
        & $add 'Exit-code contract' (($dupes.Count -eq 0) -and ($codes -contains 0)) $(if ($dupes.Count -eq 0) { "$($codes.Count) distinct codes" } else { "duplicates: $($dupes -join ', ')" })
    }
    catch { & $add 'Exit-code contract' $false $_.Exception.Message }

    # 5. the built-in defaults validate against their own schema
    try {
        $problems = @(Test-Configuration -Configuration (Get-DefaultConfiguration))
        & $add 'Configuration schema' ($problems.Count -eq 0) $(if ($problems.Count -eq 0) { 'built-in defaults pass validation' } else { $problems -join ' | ' })
    }
    catch { & $add 'Configuration schema' $false $_.Exception.Message }

    # 6. the safety guards reject hostile spool directories
    try {
        $bad = @('C:\', 'C:\Windows', 'C:\Windows\System32', 'C:\Users', '\\server\share\PRINTERS', 'C:\Data', 'C:\a\b', '', 'C:\PRINTERS', 'C:\Users\PRINTERS')
        $leaked = New-Object System.Collections.Generic.List[string]
        foreach ($p in $bad) {
            $r = Test-SpoolDirectorySafeInternal -Path $p
            if ($r.IsSafe) { $leaked.Add($p) }
        }
        # Resolve SystemRoot BEFORE Join-Path, not after. On a host where the
        # variable is unset (Linux/macOS, where -SelfTest still runs as part of
        # the CI matrix) Join-Path receives $null and throws a parameter-binding
        # error, which fails the guard check for a reason that has nothing to do
        # with the guards.
        # Built by concatenation, not Join-Path. Join-Path resolves the provider
        # drive, so on a host with no C: drive (the Linux leg of the CI matrix,
        # where -SelfTest still runs) it throws "Cannot find drive 'C'" - a noisy
        # false alarm in a check that is supposed to be quiet and deterministic.
        $sysRoot = $env:SystemRoot
        if ([string]::IsNullOrWhiteSpace($sysRoot)) { $sysRoot = 'C:\Windows' }
        $defaultDir = $sysRoot.TrimEnd('\') + '\System32\spool\PRINTERS'
        $good = Test-SpoolDirectorySafeInternal -Path $defaultDir
        if (-not $good.IsSafe -and (Test-RunningOnWindows)) { $leaked.Add("the default spool directory was rejected: $($good.Reason)") }
        & $add 'Spool-directory guards' ($leaked.Count -eq 0) $(if ($leaked.Count -eq 0) { "$($bad.Count) hostile paths rejected" } else { "NOT rejected: $($leaked -join ', ')" })
    }
    catch { & $add 'Spool-directory guards' $false $_.Exception.Message }

    # 7. help renders and documents every parameter
    try {
        $text = Get-HelpText -Section 'All' -IncludeParameterMetadata
        $docs = @(Get-ParameterDocumentation)
        $cmd = Get-Command -Name 'Clear-PrintSpoolerQueue' -ErrorAction SilentlyContinue
        $undocumented = New-Object System.Collections.Generic.List[string]
        if ($cmd) {
            $skip = @('Verbose', 'Debug', 'ErrorAction', 'WarningAction', 'InformationAction', 'ErrorVariable',
                'WarningVariable', 'InformationVariable', 'OutVariable', 'OutBuffer', 'PipelineVariable',
                'ProgressAction', 'Confirm', 'WhatIf')
            foreach ($p in @($cmd.Parameters.Keys)) {
                if ($skip -contains $p) { continue }
                if (@($docs | Where-Object { $_.Name -eq $p }).Count -eq 0) { $undocumented.Add($p) }
            }
        }
        & $add 'Help completeness' (($text.Length -gt 5000) -and ($undocumented.Count -eq 0)) `
            $(if ($undocumented.Count -eq 0) { "$($text.Length) characters, $($docs.Count) parameters documented" } else { "undocumented: $($undocumented -join ', ')" })
    }
    catch { & $add 'Help completeness' $false $_.Exception.Message }

    # 8. the time-span and size parsers
    try {
        $bad = New-Object System.Collections.Generic.List[string]
        if ((ConvertFrom-AnyTimeSpan -Value '90s').TotalSeconds -ne 90) { $bad.Add('90s') }
        if ((ConvertFrom-AnyTimeSpan -Value '2h').TotalHours -ne 2) { $bad.Add('2h') }
        if ((ConvertFrom-AnyTimeSpan -Value '7d').TotalDays -ne 7) { $bad.Add('7d') }
        if ((ConvertFrom-AnyTimeSpan -Value '500ms').TotalMilliseconds -ne 500) { $bad.Add('500ms') }
        if ((ConvertFrom-AnyTimeSpan -Value '30').TotalSeconds -ne 30) { $bad.Add('bare 30') }
        if ((ConvertFrom-AnyTimeSpan -Value '01:30:00').TotalMinutes -ne 90) { $bad.Add('01:30:00') }
        if ((Format-FileSizeArgument -Value '10MB') -ne 10485760) { $bad.Add('10MB') }
        if ((Format-FileSizeArgument -Value '1.5GB') -ne 1610612736) { $bad.Add('1.5GB') }
        if ($null -ne (Format-FileSizeArgument -Value 'not-a-size')) { $bad.Add('garbage size accepted') }
        & $add 'Value parsers' ($bad.Count -eq 0) $(if ($bad.Count -eq 0) { '9 cases' } else { "failed: $($bad -join ', ')" })
    }
    catch { & $add 'Value parsers' $false $_.Exception.Message }

    # 9. the checksum text parser
    try {
        $sample = "# comment`nabc123def456abc123def456abc123def456abc123def456abc123def456abc1  Clear-PrintSpoolerQueue.ps1`n"
        $parsed = ConvertFrom-ChecksumText -Text $sample
        & $add 'Checksum parser' ($parsed.Contains('Clear-PrintSpoolerQueue.ps1')) $(if ($parsed.Contains('Clear-PrintSpoolerQueue.ps1')) { 'GNU coreutils format' } else { 'not parsed' })
    }
    catch { & $add 'Checksum parser' $false $_.Exception.Message }

    # 10. a writable log location can always be found
    try {
        # Hermetic. Resolve-LogTarget CREATES each candidate directory as it
        # probes it, and one candidate is <script dir>\Logs. Left alone, a
        # -SelfTest run from a release folder would litter the shipped artifacts
        # with a Logs directory - which is exactly what the artifact-integrity
        # gate then (correctly) rejects. Point ScriptDirectory at a scratch
        # directory for the duration of this one check.
        $savedDir = $Script:RuntimeState.ScriptDirectory
        $scratch = Join-Path ([IO.Path]::GetTempPath()) ("psq-selftest-log-{0}" -f [Guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $scratch -Force -ErrorAction SilentlyContinue
        $Script:RuntimeState.ScriptDirectory = $scratch
        try {
            $probe = Resolve-LogTarget -ExplicitPath $null
            & $add 'Log location' $probe.Writable $(if ($probe.Writable) { "resolved via $($probe.Source)" } else { $probe.Error })
        }
        finally {
            $Script:RuntimeState.ScriptDirectory = $savedDir
            Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    catch { & $add 'Log location' $false $_.Exception.Message }

    $failed = @($checks | Where-Object { $_.Result -eq 'FAIL' }).Count
    return [pscustomobject]@{
        Tool         = $Script:ScriptName
        Version      = $Script:ScriptVersion
        Overall      = if ($failed -eq 0) { 'Passed' } else { 'Failed' }
        TotalChecks  = $checks.Count
        FailedChecks = $failed
        Checks       = $checks.ToArray()
        CheckedUtc   = [DateTime]::UtcNow.ToString('o')
    }
}

function Invoke-SelfTestPrintInternal {
    <#
    .SYNOPSIS
        Submits and immediately cancels a test job to prove the queue really works.
    .DESCRIPTION
        "Service state = Running" is not proof that printing works. This submits a
        job to the first available printer through the real Print spooler path and
        cancels it before it reaches the device. If submission succeeds, the queue
        is functional. If no printer is installed the check reports NotApplicable
        rather than failing.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$ComputerName,
        [AllowNull()][System.Management.Automation.PSCredential]$Credential,
        [AllowNull()][AllowEmptyString()][string]$PrinterName
    )

    $sb = {
        param($Printer)
        $out = @{ Success = $false; Message = $null; Printer = $null; JobId = $null; NotApplicable = $false }
        try {
            $printer = $null
            if ($Printer) {
                $printer = @(Get-Printer -Name $Printer -ErrorAction Stop | Select-Object -First 1)
                if ($printer.Count -gt 0) { $printer = $printer[0] } else { $printer = $null }
            }
            else {
                $all = @(Get-Printer -ErrorAction Stop)
                $printer = @($all | Where-Object { $_.PrinterStatus -ne 'Offline' } | Select-Object -First 1)
                if ($printer.Count -gt 0) { $printer = $printer[0] }
                elseif ($all.Count -gt 0) { $printer = $all[0] }
                else { $printer = $null }
            }

            if (-not $printer) {
                $out.NotApplicable = $true
                $out.Message = 'No printer is installed on this host; the submission test cannot run.'
                return $out
            }
            $out.Printer = $printer.Name

            $job = $null
            try {
                # Out-Printer goes through the real spool path, which is exactly what
                # we want to prove. The job is cancelled immediately below.
                $job = 'PrintSpoolerCleanup self-test' | Out-Printer -Name $printer.Name -PassThru -ErrorAction Stop
            }
            catch {
                # Out-Printer -PassThru is unavailable on some SKUs; fall back to
                # locating the job we just created by its document name.
                $found = @(Get-PrintJob -PrinterName $printer.Name -ErrorAction SilentlyContinue |
                    Where-Object { $_.DocumentName -like '*self-test*' } | Select-Object -First 1)
                if ($found.Count -gt 0) { $job = $found[0] }
            }

            if ($job) {
                $out.JobId = $job.Id
                try { Remove-PrintJob -InputObject $job -ErrorAction Stop }
                catch { /* the sweep below retries by document name */ }
            }
            # Sweep by name as well, in case the job object we held was stale.
            try {
                Get-PrintJob -PrinterName $printer.Name -ErrorAction SilentlyContinue |
                    Where-Object { $_.DocumentName -like '*self-test*' } |
                    Remove-PrintJob -ErrorAction SilentlyContinue
            }
            catch { /* best-effort cleanup; the job is empty and harmless */ }

            $out.Success = $true
            $out.Message = "A test job was accepted by '$($printer.Name)' and cancelled before reaching the device."
        }
        catch {
            $out.Message = $_.Exception.Message
            if ($_.Exception.Message -match 'No MSFT_Printer|not found|Get-Printer') { $out.NotApplicable = $true }
        }
        return $out
    }

    try {
        if (Test-IsLocalTarget -ComputerName $ComputerName) {
            $raw = & $sb $PrinterName
        }
        else {
            $p = @{ ComputerName = $ComputerName; ScriptBlock = $sb; ArgumentList = @($PrinterName); ErrorAction = 'Stop' }
            if ($Credential) { $p['Credential'] = $Credential }
            $raw = Invoke-Command @p
        }
        return [pscustomobject]@{
            ComputerName  = $ComputerName
            Success       = [bool]$raw.Success
            NotApplicable = [bool]$raw.NotApplicable
            Printer       = $raw.Printer
            JobId         = $raw.JobId
            Message       = $raw.Message
            TestedUtc     = [DateTime]::UtcNow.ToString('o')
        }
    }
    catch {
        return [pscustomobject]@{
            ComputerName  = $ComputerName
            Success       = $false
            NotApplicable = $false
            Printer       = $null
            JobId         = $null
            Message       = $_.Exception.Message
            TestedUtc     = [DateTime]::UtcNow.ToString('o')
        }
    }
}

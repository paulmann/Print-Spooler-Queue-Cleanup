# ============================================================================
#  Execution orchestration
# ----------------------------------------------------------------------------
#  Responsibilities:
#    * turn the parameter set into the flat, CLIXML-safe $Options hashtable that
#      the payload understands
#    * token-substitute the payload template ONCE per run, not per computer
#    * dispatch sequentially, in parallel (PS7) or via a runspace pool (PS5.1)
#    * pre-flight WinRM so an unreachable host fails in ~1 s instead of ~40 s
#    * normalise whatever comes back into the public result schema
#
#  Isolation guarantee: a failure on one computer can never abort the run for the
#  others. Every target produces exactly one result object.
# ============================================================================

function ConvertTo-QuotedListLiteral {
    <#
    .SYNOPSIS
        Renders a string array as a PowerShell list literal: 'a', 'b', 'c'.
    .NOTES
        Single quotes are doubled so a path containing an apostrophe
        (C:\Users\O'Brien\...) cannot break out of the literal. Top-level on
        purpose: a nested definition would be re-created on every call and would
        leak into script scope once the build inlines this file.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()][string[]]$Items)
    if (-not $Items -or @($Items).Count -eq 0) { return "''" }
    return ((@($Items) | ForEach-Object { "'{0}'" -f (([string]$_) -replace "'", "''") }) -join ', ')
}

function New-PayloadTokenMap {
    <#
    .SYNOPSIS
        Builds the @@TOKEN@@ -> literal map for the payload template.
    .NOTES
        Substitution uses [string]::Replace (ordinal), never -replace, because
        -replace is a regex operator: both the pattern and the replacement string
        would need escaping, since $ and \ are special in the replacement. A
        subtle -replace bug here would silently corrupt the payload on any site
        whose paths contain a $ (for example a hidden share such as D:\Print$).
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()

    $forbidden = @(Get-SpoolGuardForbiddenRoots)
    return @{
        '@@SPOOLER_SERVICE@@'       = $Script:SpoolerServiceName
        '@@DEPENDENT_SERVICES@@'    = (ConvertTo-QuotedListLiteral -Items $Script:DependentServices)
        '@@SPOOL_EXTENSIONS@@'      = (ConvertTo-QuotedListLiteral -Items $Script:SpoolExtensions)
        '@@LEGACY_EXTENSIONS@@'     = (ConvertTo-QuotedListLiteral -Items $Script:LegacySpoolExtensions)
        '@@DEFAULT_SPOOL_SUBPATH@@' = $Script:DefaultSpoolDirectory
        '@@FORMS_SUBPATH@@'         = 'System32\spool\PRINTERS\..\..\..\..\System32\spool\PRINTERS'
        '@@GUARD_REQUIRED_LEAF@@'   = $Script:SpoolGuardRequiredLeaf
        '@@GUARD_MIN_DEPTH@@'       = [string]$Script:SpoolGuardMinPathDepth
        '@@GUARD_MAX_FILES@@'       = [string]$Script:SpoolGuardMaxFileCount
        '@@GUARD_MAX_BYTES@@'       = [string][long]$Script:SpoolGuardMaxTotalBytes
        '@@FORBIDDEN_ROOTS@@'       = (ConvertTo-QuotedListLiteral -Items $forbidden)
        '@@RETRY_BACKOFF@@'         = [string]$Script:RetryBackoffMultiplier
        '@@SCRIPT_NAME@@'           = $Script:ScriptName
        '@@SCRIPT_VERSION@@'        = $Script:ScriptVersion
        '@@PRODUCT_FAMILY@@'        = $Script:ProductFamily
    }
}

function Resolve-PayloadScriptBlock {
    <#
    .SYNOPSIS
        Substitutes the tokens and returns a ready-to-invoke [scriptblock].
    .DESCRIPTION
        Cached for the whole run: substitution is deterministic, so doing it once
        per computer would be wasted work and a consistency risk. The generated
        text is parsed here so a template defect fails immediately and loudly,
        rather than at 03:00 on a print server in the middle of a maintenance window.
    #>
    [CmdletBinding()]
    [OutputType([scriptblock])]
    param([switch]$Force)

    if ($Script:RuntimeState.PayloadScriptBlock -and -not $Force) { return $Script:RuntimeState.PayloadScriptBlock }

    $text = $Script:RemotePayloadTemplate
    $map = New-PayloadTokenMap
    foreach ($token in @($map.Keys)) {
        $text = $text.Replace($token, [string]$map[$token])
    }

    $tokens = $null
    $errors = $null
    $null = [System.Management.Automation.Language.Parser]::ParseInput($text, [ref]$tokens, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) {
        $detail = ($errors | Select-Object -First 5 | ForEach-Object { "line $($_.Extent.StartLineNumber): $($_.Message)" }) -join '; '
        throw "Internal error: the generated cleanup payload does not parse. $detail"
    }
    $leftover = @([regex]::Matches($text, '@@[A-Z_]+@@') | ForEach-Object { $_.Value } | Select-Object -Unique)
    if ($leftover.Count -gt 0) {
        throw "Internal error: unsubstituted payload token(s): $($leftover -join ', ')"
    }

    $Script:RuntimeState.PayloadScriptBlock = [scriptblock]::Create($text)
    $Script:RuntimeState.PayloadText = $text
    return $Script:RuntimeState.PayloadScriptBlock
}

function Get-PayloadValue {
    <#
    .SYNOPSIS
        Null-safe dictionary reader used while normalising payload results.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowNull()][System.Collections.IDictionary]$Hashtable,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Key,
        [AllowNull()]$Default
    )
    if ($null -eq $Hashtable) { return $Default }
    if (-not $Hashtable.Contains($Key)) { return $Default }
    $v = $Hashtable[$Key]
    if ($null -eq $v) { return $Default }
    return $v
}

function New-PayloadOptions {
    <#
    .SYNOPSIS
        Flattens the resolved configuration and parameters into the payload contract.
    .NOTES
        Only primitives, strings and string arrays. Everything must survive CLIXML
        serialization to a Windows PowerShell 5.1 remote endpoint.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)][AllowNull()][System.Collections.IDictionary]$Config,
        [Parameter(Mandatory)][AllowNull()][System.Collections.IDictionary]$Arguments
    )

    $olderThan = 0.0
    if ($Arguments['OlderThan']) { $olderThan = (ConvertFrom-AnyTimeSpan -Value $Arguments['OlderThan']).TotalSeconds }
    $newerThan = 0.0
    if ($Arguments['NewerThan']) { $newerThan = (ConvertFrom-AnyTimeSpan -Value $Arguments['NewerThan']).TotalSeconds }

    return @{
        Mode                     = [string]$Arguments['Mode']
        DryRun                   = [bool]$Arguments['DryRun']
        Repair                   = [bool]$Arguments['Repair']
        Backup                   = [bool]$Arguments['Backup']
        BackupRequired           = [bool]$Arguments['BackupRequired']
        BackupPath               = [string]$Arguments['BackupPath']
        SpoolDirectoryOverride   = [string]$Arguments['SpoolDirectory']

        ServiceTimeoutSeconds    = [int]$Config['Service']['TimeoutSeconds']
        SettleSeconds            = [int]$Config['Service']['SettleSeconds']
        StopDependents           = [bool]$Config['Service']['StopDependentServices']
        RestartDependents        = [bool]$Config['Service']['RestartDependentServices']
        ForceStopFallback        = [bool]$Config['Service']['ForceStopFallback']

        RetryCount               = [int]$Config['Cleanup']['RetryCount']
        RetryDelayMs             = [int]$Config['Cleanup']['RetryDelayMs']
        QuiescenceSeconds        = [int]$Config['Cleanup']['WaitForQuiescenceSeconds']
        VerifyAfterRestart       = [bool]$Config['Cleanup']['VerifyAfterRestart']
        PurgeLegacy              = [bool]$Config['Cleanup']['PurgeLegacyExtensions']

        PrinterName              = [string[]]@($Arguments['PrinterName'])
        ExcludePrinterName       = [string[]]@($Arguments['ExcludePrinterName'])
        JobId                    = [long[]]@($Arguments['JobId'])
        DocumentName             = [string]$Arguments['DocumentName']
        OwnerName                = [string]$Arguments['OwnerName']
        OlderThanSeconds         = [double]$olderThan
        NewerThanSeconds         = [double]$newerThan
        LargerThanBytes          = [long]$Arguments['LargerThanBytes']
        SmallerThanBytes         = [long]$Arguments['SmallerThanBytes']

        PurgeFormsCache          = [bool]$Arguments['PurgeFormsCache']
        PurgeDriverStore         = [bool]$Arguments['PurgeDriverStore']
        ResetSpoolDirectoryValue = [bool]$Arguments['ResetSpoolDirectoryValue']

        EventLogHours            = [int]$Config['Diagnostics']['EventLogHours']
        EventLogMaxEntries       = [int]$Config['Diagnostics']['EventLogMaxEntries']
    }
}

function Test-WinRmAvailability {
    <#
    .SYNOPSIS
        Fast reachability and authorisation probe for one remote target.
    .DESCRIPTION
        FINDING A11: without a probe, every unreachable host costs a full WinRM
        timeout (~40 s) before failing. Probing first means a fleet of 200 hosts
        with 30 offline reports the offline ones in seconds and never delays the
        healthy ones. The probe runs inside a job so a hung WinRM stack cannot
        block the caller beyond the timeout.
    .OUTPUTS
        PSCustomObject { ComputerName, Reachable, Reason, ElapsedMs }
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$ComputerName,
        [int]$TimeoutSeconds = 10,
        [AllowNull()][System.Management.Automation.PSCredential]$Credential,
        [switch]$UseSsl
    )

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $result = [pscustomobject]@{
        ComputerName = $ComputerName
        Reachable    = $false
        Reason       = $null
        ElapsedMs    = [long]0
    }

    $job = $null
    try {
        $params = @{
            ComputerName = $ComputerName
            ScriptBlock  = { $true }
            ErrorAction  = 'Stop'
        }
        if ($Credential) { $params['Credential'] = $Credential }
        if ($UseSsl) { $params['UseSsl'] = $true }

        $job = Start-Job -ScriptBlock {
            param($p)
            try { Invoke-Command @p | Out-Null; 'OK' }
            catch { $_.Exception.Message }
        } -ArgumentList $params

        if (Wait-Job -Job $job -Timeout $TimeoutSeconds) {
            $outcome = Receive-Job -Job $job -ErrorAction SilentlyContinue
            if ($outcome -eq 'OK') { $result.Reachable = $true }
            else { $result.Reason = [string](@($outcome)[0]) }
        }
        else {
            $result.Reason = "No WinRM response within ${TimeoutSeconds}s (host offline, WinRM disabled, or a firewall blocking 5985/5986)."
        }
    }
    catch {
        $result.Reason = $_.Exception.Message
    }
    finally {
        if ($job) { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue }
    }

    $sw.Stop()
    $result.ElapsedMs = $sw.ElapsedMilliseconds
    return $result
}

function ConvertFrom-PayloadResult {
    <#
    .SYNOPSIS
        Maps the payload hashtable onto the public result object (schema contract).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowNull()]$Raw,
        [Parameter(Mandatory)][AllowEmptyString()][string]$ComputerName
    )

    $r = New-ResultObject -ComputerName $ComputerName
    $r.IsLocal = Test-IsLocalTarget -ComputerName $ComputerName
    if (-not $Raw) {
        $r.Status = 'Failed'
        $r.FinishedUtc = [DateTime]::UtcNow
        return $r
    }

    $r.Status               = [string](Get-PayloadValue -Hashtable $Raw -Key 'Status' -Default 'Unknown')
    $r.SpoolDirectory       = Get-PayloadValue -Hashtable $Raw -Key 'SpoolDirectory' -Default $null
    $r.SpoolDirectorySource = Get-PayloadValue -Hashtable $Raw -Key 'SpoolDirectorySource' -Default $null
    $r.ServiceWasRunning    = Get-PayloadValue -Hashtable $Raw -Key 'ServiceWasRunning' -Default $null
    $r.ServiceIsRunning     = Get-PayloadValue -Hashtable $Raw -Key 'ServiceIsRunning' -Default $null
    $r.ServiceStopMs        = [long](Get-PayloadValue -Hashtable $Raw -Key 'ServiceStopMs' -Default 0)
    $r.ServiceStartMs       = [long](Get-PayloadValue -Hashtable $Raw -Key 'ServiceStartMs' -Default 0)
    $r.DurationMs           = [long](Get-PayloadValue -Hashtable $Raw -Key 'DurationMs' -Default 0)
    $r.JobsFound            = [int](Get-PayloadValue -Hashtable $Raw -Key 'JobsFound' -Default 0)
    $r.JobsSelected         = [int](Get-PayloadValue -Hashtable $Raw -Key 'JobsSelected' -Default 0)
    $r.JobsRemoved          = [int](Get-PayloadValue -Hashtable $Raw -Key 'JobsRemoved' -Default 0)
    $r.FilesFound           = [int](Get-PayloadValue -Hashtable $Raw -Key 'FilesFound' -Default 0)
    $r.FilesSelected        = [int](Get-PayloadValue -Hashtable $Raw -Key 'FilesSelected' -Default 0)
    $r.FilesDeleted         = [int](Get-PayloadValue -Hashtable $Raw -Key 'FilesDeleted' -Default 0)
    $r.FilesFailed          = [int](Get-PayloadValue -Hashtable $Raw -Key 'FilesFailed' -Default 0)
    $r.FilesLocked          = [int](Get-PayloadValue -Hashtable $Raw -Key 'FilesLocked' -Default 0)
    $r.BytesFound           = [long](Get-PayloadValue -Hashtable $Raw -Key 'BytesFound' -Default 0)
    $r.BytesDeleted         = [long](Get-PayloadValue -Hashtable $Raw -Key 'BytesDeleted' -Default 0)
    $r.FilesAfter           = [int](Get-PayloadValue -Hashtable $Raw -Key 'FilesAfter' -Default 0)
    $r.ResurrectedFiles     = [int](Get-PayloadValue -Hashtable $Raw -Key 'ResurrectedFiles' -Default 0)
    $r.BackupPath           = Get-PayloadValue -Hashtable $Raw -Key 'BackupPath' -Default $null
    $r.BackupHash           = Get-PayloadValue -Hashtable $Raw -Key 'BackupHash' -Default $null
    $r.DryRun               = [bool](Get-PayloadValue -Hashtable $Raw -Key 'DryRun' -Default $false)
    $r.RepairPerformed      = [bool](Get-PayloadValue -Hashtable $Raw -Key 'RepairPerformed' -Default $false)
    $r.DiagnosticsCollected = [bool](Get-PayloadValue -Hashtable $Raw -Key 'DiagnosticsCollected' -Default $false)
    $r.Mode                 = Get-PayloadValue -Hashtable $Raw -Key 'Mode' -Default $null

    $r.DependentServices = @($Raw['DependentServices'])
    $r.Jobs              = @($Raw['Jobs'])
    $r.RemovedFiles      = @($Raw['RemovedFiles'])
    $r.FailedFiles       = @($Raw['FailedFiles'])
    $r.OrphanedFiles     = @($Raw['OrphanedFiles'])
    $r.ServiceTrace      = @($Raw['ServiceTrace'])
    $r.Diagnostics       = $Raw['Diagnostics']

    foreach ($w in @($Raw['Warnings'])) { if ($w) { $r.Warnings.Add([string]$w) | Out-Null } }
    foreach ($e in @($Raw['Errors'])) { if ($e) { $r.Errors.Add([string]$e) | Out-Null } }

    $started = Get-PayloadValue -Hashtable $Raw -Key 'StartedUtc' -Default $null
    $finished = Get-PayloadValue -Hashtable $Raw -Key 'FinishedUtc' -Default $null
    try {
        if ($started) { $r.StartedUtc = [DateTime]::Parse([string]$started, $null, [Globalization.DateTimeStyles]::RoundtripKind) }
    }
    catch { $r.StartedUtc = [DateTime]::UtcNow }
    try {
        if ($finished) { $r.FinishedUtc = [DateTime]::Parse([string]$finished, $null, [Globalization.DateTimeStyles]::RoundtripKind) }
        else { $r.FinishedUtc = [DateTime]::UtcNow }
    }
    catch { $r.FinishedUtc = [DateTime]::UtcNow }

    return $r
}

function Invoke-CleanupOnTarget {
    <#
    .SYNOPSIS
        Runs the payload against one computer. Local means a direct invoke.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$ComputerName,
        [Parameter(Mandatory)][AllowNull()][System.Collections.IDictionary]$Options,
        [Parameter(Mandatory)][scriptblock]$Payload,
        [AllowNull()][System.Management.Automation.PSCredential]$Credential,
        [switch]$UseSsl
    )

    if (Test-IsLocalTarget -ComputerName $ComputerName) {
        $raw = & $Payload $Options
    }
    else {
        $params = @{
            ComputerName = $ComputerName
            ScriptBlock  = $Payload
            ArgumentList = @(, $Options)
            ErrorAction  = 'Stop'
        }
        if ($Credential) { $params['Credential'] = $Credential }
        if ($UseSsl) { $params['UseSsl'] = $true }
        $raw = Invoke-Command @params
    }

    # The payload always returns ONE hashtable, but a stray Write-Output anywhere
    # upstream would wrap it in an array. Unwrap defensively.
    if (($raw -is [System.Collections.IEnumerable]) -and ($raw -isnot [string]) -and ($raw -isnot [System.Collections.IDictionary])) {
        $raw = @($raw | Where-Object { $_ -is [System.Collections.IDictionary] } | Select-Object -First 1)
    }
    if (-not $raw) { throw "No result was returned for '$ComputerName'." }
    return $raw
}

function Invoke-CleanupFleet {
    <#
    .SYNOPSIS
        Executes the payload across the target list: sequential, PS7 -Parallel,
        or a PS5.1 runspace pool.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowNull()][AllowEmptyCollection()][string[]]$Targets,
        [Parameter(Mandatory)][AllowNull()][System.Collections.IDictionary]$Options,
        [Parameter(Mandatory)][scriptblock]$Payload,
        [AllowNull()][System.Management.Automation.PSCredential]$Credential,
        [int]$ThrottleLimit = 1,
        [switch]$UseSsl,
        [bool]$PreFlightProbe = $true,
        [int]$ProbeTimeoutSeconds = 10
    )

    $results = New-Object System.Collections.Generic.List[object]
    if (-not $Targets -or @($Targets).Count -eq 0) { return $results.ToArray() }

    $local = New-Object System.Collections.Generic.List[string]
    $remote = New-Object System.Collections.Generic.List[string]
    foreach ($t in $Targets) {
        if ([string]::IsNullOrWhiteSpace($t)) { continue }
        if (Test-IsLocalTarget -ComputerName $t) { $local.Add($t) }
        else { $remote.Add($t) }
    }

    # ---- 1. local first, always sequential (it is exactly one machine) -------
    foreach ($t in $local.ToArray()) {
        Write-LogInfo "Processing the local target '$t'..." -Component 'Dispatch' -Target $t
        Show-CleanupProgress -Activity 'Print Spooler Cleanup' -Status "Local: $t" -PercentComplete 5
        try {
            $raw = Invoke-CleanupOnTarget -ComputerName $t -Options $Options -Payload $Payload -UseSsl:$UseSsl
            $results.Add((ConvertFrom-PayloadResult -Raw $raw -ComputerName $t))
        }
        catch {
            Write-LogError "Local cleanup failed: $($_.Exception.Message)" -Component 'Dispatch' -Target $t
            $r = New-ResultObject -ComputerName $t -Status 'Failed'
            $r.IsLocal = $true
            $r.Errors.Add($_.Exception.Message) | Out-Null
            $r.FinishedUtc = [DateTime]::UtcNow
            $results.Add($r)
        }
    }
    if ($remote.Count -eq 0) { return $results.ToArray() }

    # ---- 2. optional WinRM pre-flight probe ---------------------------------
    $reachable = New-Object System.Collections.Generic.List[string]
    if ($PreFlightProbe) {
        Write-LogInfo "Probing WinRM on $($remote.Count) remote target(s)..." -Component 'Remote'
        Show-CleanupProgress -Activity 'Print Spooler Cleanup' -Status 'Probing WinRM' -PercentComplete 10
        foreach ($t in $remote.ToArray()) {
            $probe = Test-WinRmAvailability -ComputerName $t -TimeoutSeconds $ProbeTimeoutSeconds -Credential $Credential -UseSsl:$UseSsl
            if ($probe.Reachable) {
                $reachable.Add($t)
                Write-LogDebug "WinRM probe OK for '$t' in $($probe.ElapsedMs) ms." -Component 'Remote' -Target $t
            }
            else {
                Write-LogWarn "WinRM probe failed for '$t': $($probe.Reason)" -Component 'Remote' -Target $t
                $r = New-ResultObject -ComputerName $t -Status 'Failed'
                $r.Errors.Add("WinRM pre-flight failed: $($probe.Reason)") | Out-Null
                $r.FinishedUtc = [DateTime]::UtcNow
                $results.Add($r)
            }
        }
    }
    else {
        foreach ($t in $remote.ToArray()) { $reachable.Add($t) }
    }
    if ($reachable.Count -eq 0) { return $results.ToArray() }

    $effectiveThrottle = [Math]::Max(1, [Math]::Min($ThrottleLimit, $reachable.Count, $Script:MaxThrottleLimit))
    $payloadText = $Payload.ToString()

    # ---- 3a. PowerShell 7: ForEach-Object -Parallel -------------------------
    if ($effectiveThrottle -gt 1 -and (Test-PowerShellSevenFeatures)) {
        Write-LogInfo "Dispatching $($reachable.Count) remote target(s) in parallel (throttle $effectiveThrottle)..." -Component 'Remote'
        # NOTE: no Write-Progress inside the worker. Progress written from a
        # -Parallel runspace is not reliably surfaced in the parent host, and
        # $using: can only capture variables that exist in the CALLER scope -
        # referencing an undefined one is a parse-time error. Overall progress is
        # rendered from the collection loop below instead.
        $sslFlag = [bool]$UseSsl
        $totalRemote = $reachable.Count
        $parallel = $reachable.ToArray() | ForEach-Object -ThrottleLimit $effectiveThrottle -Parallel {
            $computer = $_
            $opts = $using:Options
            $cred = $using:Credential
            $ssl = $using:sslFlag
            $text = $using:payloadText
            try {
                $sb = [scriptblock]::Create($text)
                $p = @{ ComputerName = $computer; ScriptBlock = $sb; ArgumentList = @(, $opts); ErrorAction = 'Stop' }
                if ($cred) { $p['Credential'] = $cred }
                if ($ssl) { $p['UseSsl'] = $true }
                [pscustomobject]@{ Computer = $computer; Raw = (Invoke-Command @p); Error = $null }
            }
            catch {
                [pscustomobject]@{ Computer = $computer; Raw = $null; Error = $_.Exception.Message }
            }
        }

        $seen = 0
        foreach ($item in $parallel) {
            $seen++
            Show-CleanupProgress -Activity 'Print Spooler Cleanup' -Status "Remote $seen / $totalRemote hosts completed" -PercentComplete ([int](($seen / [Math]::Max(1, $totalRemote)) * 100)) -Id 2
            if ($item.Error) {
                Write-LogError "Remote cleanup failed on '$($item.Computer)': $($item.Error)" -Component 'Remote' -Target $item.Computer
                $r = New-ResultObject -ComputerName $item.Computer -Status 'Failed'
                $r.Errors.Add($item.Error) | Out-Null
                $r.FinishedUtc = [DateTime]::UtcNow
                $results.Add($r)
            }
            else {
                $results.Add((ConvertFrom-PayloadResult -Raw $item.Raw -ComputerName $item.Computer))
            }
        }
        Show-CleanupProgress -Activity 'Print Spooler Cleanup' -Id 2 -Completed
        return $results.ToArray()
    }

    # ---- 3b. Windows PowerShell 5.1: runspace pool --------------------------
    if ($effectiveThrottle -gt 1) {
        Write-LogInfo "Dispatching $($reachable.Count) remote target(s) via a runspace pool (throttle $effectiveThrottle)..." -Component 'Remote'
        $pool = [runspacefactory]::CreateRunspacePool(1, $effectiveThrottle)
        $pool.Open()
        $workers = New-Object System.Collections.Generic.List[object]
        try {
            foreach ($t in $remote.ToArray()) {
                if ($reachable -notcontains $t) { continue }
                $ps = [PowerShell]::Create()
                $ps.RunspacePool = $pool
                [void]$ps.AddScript({
                    param($computer, $opts, $text, $cred, $ssl)
                    try {
                        $sb = [scriptblock]::Create($text)
                        $p = @{ ComputerName = $computer; ScriptBlock = $sb; ArgumentList = @(, $opts); ErrorAction = 'Stop' }
                        if ($cred) { $p['Credential'] = $cred }
                        if ($ssl) { $p['UseSsl'] = $true }
                        [pscustomobject]@{ Computer = $computer; Raw = (Invoke-Command @p); Error = $null }
                    }
                    catch {
                        [pscustomobject]@{ Computer = $computer; Raw = $null; Error = $_.Exception.Message }
                    }
                }).AddArgument($t).AddArgument($Options).AddArgument($payloadText).AddArgument($Credential).AddArgument([bool]$UseSsl)
                $workers.Add([pscustomobject]@{ PowerShell = $ps; Handle = $ps.BeginInvoke(); Computer = $t })
            }

            $done = 0
            foreach ($w in $workers) {
                try {
                    $out = @($w.PowerShell.EndInvoke($w.Handle))
                    $item = if ($out.Count -gt 0) { $out[0] } else { $null }
                    if ($item -and $item.Error) {
                        Write-LogError "Remote cleanup failed on '$($item.Computer)': $($item.Error)" -Component 'Remote' -Target $item.Computer
                        $r = New-ResultObject -ComputerName $item.Computer -Status 'Failed'
                        $r.Errors.Add($item.Error) | Out-Null
                        $r.FinishedUtc = [DateTime]::UtcNow
                        $results.Add($r)
                    }
                    elseif ($item) {
                        $results.Add((ConvertFrom-PayloadResult -Raw $item.Raw -ComputerName $item.Computer))
                    }
                }
                catch {
                    Write-LogError "Remote worker error for '$($w.Computer)': $($_.Exception.Message)" -Component 'Remote' -Target $w.Computer
                    $r = New-ResultObject -ComputerName $w.Computer -Status 'Failed'
                    $r.Errors.Add($_.Exception.Message) | Out-Null
                    $r.FinishedUtc = [DateTime]::UtcNow
                    $results.Add($r)
                }
                finally {
                    $done++
                    Show-CleanupProgress -Activity 'Print Spooler Cleanup' -Status "$done / $($workers.Count) remote hosts" -PercentComplete ([int](($done / [Math]::Max(1, $workers.Count)) * 100)) -Id 2
                }
            }
        }
        finally {
            foreach ($w in $workers) { try { $w.PowerShell.Dispose() } catch { /* the pool is closing anyway */ } }
            $pool.Close()
            $pool.Dispose()
        }
        Show-CleanupProgress -Activity 'Print Spooler Cleanup' -Id 2 -Completed
        return $results.ToArray()
    }

    # ---- 3c. sequential remote ----------------------------------------------
    Write-LogInfo "Dispatching $($reachable.Count) remote target(s) sequentially..." -Component 'Remote'
    $i = 0
    foreach ($t in $reachable.ToArray()) {
        $i++
        Show-CleanupProgress -Activity 'Print Spooler Cleanup' -Status "Remote $i / $($reachable.Count): $t" -PercentComplete ([int](($i / $reachable.Count) * 100))
        Write-LogInfo "Processing the remote target '$t' ($i/$($reachable.Count))..." -Component 'Remote' -Target $t
        try {
            $raw = Invoke-CleanupOnTarget -ComputerName $t -Options $Options -Payload $Payload -Credential $Credential -UseSsl:$UseSsl
            $results.Add((ConvertFrom-PayloadResult -Raw $raw -ComputerName $t))
        }
        catch {
            Write-LogError "Remote cleanup failed on '$t': $($_.Exception.Message)" -Component 'Remote' -Target $t
            $r = New-ResultObject -ComputerName $t -Status 'Failed'
            $r.Errors.Add($_.Exception.Message) | Out-Null
            $r.FinishedUtc = [DateTime]::UtcNow
            $results.Add($r)
        }
    }
    Show-CleanupProgress -Activity 'Print Spooler Cleanup' -Completed
    return $results.ToArray()
}

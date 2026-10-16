# ============================================================================
#  Task Scheduler integration (finding A7)
# ----------------------------------------------------------------------------
#  v3 documented a schtasks one-liner in the README and nothing else. That
#  one-liner was wrong in two ways: it ran the interactive .bat as SYSTEM (so
#  "Press any key" blocked the task until its time limit killed it) and it set no
#  execution time limit. v4 ships a real installer that:
#    * always launches the NON-interactive PowerShell entry point with -Force
#      -NonInteractive -Quiet, never the .bat
#    * supports daily / weekly / hourly / at-startup / on-event schedules
#    * supports a computer-list file so one task can service a whole site
#    * sets an execution time limit, disables the "stop on battery" defaults that
#      silently skip maintenance on laptops, and enables restart-on-failure
#    * registers either through the ScheduledTasks module (preferred) or through a
#      UTF-16LE XML file plus schtasks.exe (works on Server Core and 2012)
# ============================================================================

function New-SpoolerTaskXml {
    <#
    .SYNOPSIS
        Builds a Task Scheduler 1.4 XML definition as a string.
    .NOTES
        Returned as a plain string; the caller writes it as UTF-16LE when using
        `schtasks /create /xml`, because schtasks requires that encoding.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Command,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Arguments,
        [string]$WorkingDirectory = 'C:\Windows\System32',
        [string]$RunAsAccount = 'NT AUTHORITY\SYSTEM',
        [ValidateSet('Daily', 'Weekly', 'Hourly', 'AtStartup', 'OnEvent', 'Once')][string]$Schedule = 'Daily',
        [string]$StartTime = '03:00',
        [int]$IntervalHours = 24,
        [AllowNull()][AllowEmptyString()][string]$EventLogName,
        [AllowNull()][AllowEmptyString()][string]$EventSource,
        [AllowNull()][Nullable[int]]$EventId,
        [string]$ExecutionTimeLimit = 'PT30M',
        [string]$Description = 'Automated Windows Print Spooler queue maintenance.',
        [string]$Author = 'Print Spooler Queue Cleanup',
        [switch]$RunOnlyIfIdle,
        [int]$RandomDelayMinutes = 15
    )

    $startBoundary = (Get-Date -Hour 0 -Minute 0 -Second 0).AddDays(1)
    try {
        $parts = $StartTime.Split(':')
        $startBoundary = Get-Date -Hour ([int]$parts[0]) -Minute ([int]$parts[1]) -Second 0
        if (($startBoundary -lt (Get-Date)) -and ($Schedule -ne 'AtStartup')) { $startBoundary = $startBoundary.AddDays(1) }
    }
    catch {
        # An unparseable StartTime keeps the safe next-midnight default.
    }
    $boundaryIso = $startBoundary.ToString('yyyy-MM-ddTHH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
    $delayMinutes = [Math]::Max(1, $RandomDelayMinutes)
    $interval = [Math]::Max(1, $IntervalHours)

    $trigger = switch ($Schedule) {
        'Daily' {
            @"
    <CalendarTrigger>
      <StartBoundary>$boundaryIso</StartBoundary>
      <Enabled>true</Enabled>
      <RandomDelay>PT${delayMinutes}M</RandomDelay>
      <ScheduleByDay><DaysInterval>1</DaysInterval></ScheduleByDay>
    </CalendarTrigger>
"@
        }
        'Weekly' {
            @"
    <CalendarTrigger>
      <StartBoundary>$boundaryIso</StartBoundary>
      <Enabled>true</Enabled>
      <RandomDelay>PT${delayMinutes}M</RandomDelay>
      <ScheduleByWeek>
        <DaysOfWeek><Monday/><Tuesday/><Wednesday/><Thursday/><Friday/></DaysOfWeek>
        <WeeksInterval>1</WeeksInterval>
      </ScheduleByWeek>
    </CalendarTrigger>
"@
        }
        'Hourly' {
            @"
    <CalendarTrigger>
      <StartBoundary>$boundaryIso</StartBoundary>
      <Enabled>true</Enabled>
      <Repetition>
        <Interval>PT${interval}H</Interval>
        <Duration>P1D</Duration>
        <StopAtDurationEnd>false</StopAtDurationEnd>
      </Repetition>
      <ScheduleByDay><DaysInterval>1</DaysInterval></ScheduleByDay>
    </CalendarTrigger>
"@
        }
        'AtStartup' {
            @"
    <BootTrigger>
      <Enabled>true</Enabled>
      <Delay>PT2M</Delay>
    </BootTrigger>
"@
        }
        'OnEvent' {
            $logEsc = [System.Security.SecurityElement]::Escape($EventLogName)
            $srcEsc = [System.Security.SecurityElement]::Escape($EventSource)
            $idPart = if ($EventId) { " and EventID=$EventId" } else { '' }
            $sub = "<QueryList><Query Id=""0"" Path=""$logEsc""><Select Path=""$logEsc"">*[System[Provider[@Name='$srcEsc']$idPart]]</Select></Query></QueryList>"
            @"
    <EventTrigger>
      <Enabled>true</Enabled>
      <Subscription>$sub</Subscription>
      <Delay>PT1M</Delay>
    </EventTrigger>
"@
        }
        default {
            @"
    <TimeTrigger>
      <StartBoundary>$boundaryIso</StartBoundary>
      <Enabled>true</Enabled>
    </TimeTrigger>
"@
        }
    }

    if ($RunOnlyIfIdle) {
        $idle = @"
    <IdleSettings>
      <Duration>PT10M</Duration>
      <WaitTimeout>PT1H</WaitTimeout>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <RunOnlyIfIdle>true</RunOnlyIfIdle>
"@
    }
    else {
        $idle = @"
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
"@
    }

    $escCommand = [System.Security.SecurityElement]::Escape($Command)
    $escArgs = [System.Security.SecurityElement]::Escape($Arguments)
    $escWorkDir = [System.Security.SecurityElement]::Escape($WorkingDirectory)
    $escDesc = [System.Security.SecurityElement]::Escape($Description)
    $escAuthor = [System.Security.SecurityElement]::Escape($Author)
    $escRunAs = [System.Security.SecurityElement]::Escape($RunAsAccount)
    $escLimit = [System.Security.SecurityElement]::Escape($ExecutionTimeLimit)
    $nowIso = [DateTime]::Now.ToString('yyyy-MM-ddTHH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)

    return @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>$nowIso</Date>
    <Author>$escAuthor</Author>
    <Description>$escDesc</Description>
  </RegistrationInfo>
  <Triggers>
$trigger
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>$escRunAs</UserId>
      <RunLevel>HighestAvailable</RunLevel>
      <LogonType>ServiceAccount</LogonType>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <ExecutionTimeLimit>$escLimit</ExecutionTimeLimit>
    <Priority>7</Priority>
    <RestartOnFailure>
      <Interval>PT10M</Interval>
      <Count>2</Count>
    </RestartOnFailure>
$idle
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>$escCommand</Command>
      <Arguments>$escArgs</Arguments>
      <WorkingDirectory>$escWorkDir</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"@
}

function New-SpoolerTaskArgumentString {
    <#
    .SYNOPSIS
        Builds the PowerShell argument string the scheduled task will execute.
    .DESCRIPTION
        Every switch that could block an unattended run is disabled explicitly.
        This is the precise failure mode of the schtasks one-liner that v3
        documented: the task started, hit a confirmation prompt nobody could
        answer, and sat there until the time limit killed it.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$ScriptPath,
        [AllowNull()][AllowEmptyCollection()][string[]]$ComputerName,
        [AllowNull()][AllowEmptyString()][string]$ComputerListFile,
        [AllowNull()][AllowEmptyString()][string]$ExtraArguments,
        [switch]$Backup,
        [switch]$AutoUpdate
    )

    $argList = New-Object System.Collections.Generic.List[string]
    $argList.Add('-NoProfile')
    $argList.Add('-ExecutionPolicy')
    $argList.Add('Bypass')
    $argList.Add('-File')
    $argList.Add(('"{0}"' -f $ScriptPath))
    $argList.Add('-Force')
    $argList.Add('-NonInteractive')
    $argList.Add('-Quiet')
    $argList.Add('-NoProgress')
    if ($Backup) { $argList.Add('-Backup') }
    if ($AutoUpdate) { $argList.Add('-AutoUpdate') }

    if ($ComputerName -and @($ComputerName).Count -gt 0) {
        $argList.Add('-ComputerName')
        $argList.Add(('"{0}"' -f (@($ComputerName) -join '","')))
    }
    elseif (-not [string]::IsNullOrWhiteSpace($ComputerListFile)) {
        $argList.Add('-ComputerListFile')
        $argList.Add(('"{0}"' -f $ComputerListFile))
    }
    if (-not [string]::IsNullOrWhiteSpace($ExtraArguments)) { $argList.Add($ExtraArguments) }

    return ($argList -join ' ')
}

function Get-SpoolerTaskFolderAndLeaf {
    <#
    .SYNOPSIS
        Splits 'Folder\Sub\TaskName' into a TaskPath and a leaf name.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$TaskName)

    $folder = '\'
    $leaf = $TaskName
    if ($TaskName -match '^(?<f>.+)[\\/](?<l>[^\\/]+)$') {
        $folder = '\' + $Matches['f'].Trim('\', '/') + '\'
        $leaf = $Matches['l']
    }
    return [pscustomobject]@{ Folder = $folder; Leaf = $leaf }
}

function New-SpoolerTaskFolder {
    <#
    .SYNOPSIS
        Creates each missing Task Scheduler folder segment in order.
    .NOTES
        Register-ScheduledTask cannot create folders, so a nested task path fails
        without this. Uses the Schedule.Service COM API, which is present on every
        supported SKU including Server Core.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Folder)

    if ([string]::IsNullOrWhiteSpace($Folder) -or $Folder -eq '\') { return $true }
    $svc = $null
    try {
        $svc = New-Object -ComObject Schedule.Service
        $svc.Connect()
        $segments = @($Folder.Trim('\').Split('\') | Where-Object { $_ })
        $current = ''
        foreach ($segment in $segments) {
            $parentPath = if ($current) { $current } else { '\' }
            $parent = $svc.GetFolder($parentPath)
            $current = if ($current) { "$current\$segment" } else { "\$segment" }
            $exists = $true
            try { $null = $svc.GetFolder($current) }
            catch { $exists = $false }
            if (-not $exists) { $null = $parent.CreateFolder($segment) }
        }
        return $true
    }
    catch {
        Write-LogWarn "Could not create the task folder '$Folder' ($($_.Exception.Message)); registration will fall back to the XML method." -Component 'Task'
        return $false
    }
    finally {
        if ($svc) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($svc) }
    }
}

function Install-SpoolerCleanupTaskInternal {
    <#
    .SYNOPSIS
        Registers or replaces the maintenance task.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$TaskName,
        [Parameter(Mandatory)][AllowEmptyString()][string]$ScriptPath,
        [string]$Schedule = 'Daily',
        [string]$StartTime = '03:00',
        [int]$IntervalHours = 6,
        [AllowNull()][AllowEmptyCollection()][string[]]$ComputerName,
        [AllowNull()][AllowEmptyString()][string]$ComputerListFile,
        [AllowNull()][AllowEmptyString()][string]$Arguments,
        [string]$RunAsAccount = 'NT AUTHORITY\SYSTEM',
        [AllowNull()][AllowEmptyString()][string]$EventLogName,
        [AllowNull()][AllowEmptyString()][string]$EventSource,
        [AllowNull()][Nullable[int]]$EventId,
        [int]$ExecutionTimeLimitMinutes = 30,
        [switch]$RunOnlyIfIdle,
        [switch]$Backup,
        [switch]$AutoUpdate,
        [switch]$PassThru,
        [ValidateSet('Auto', 'Module', 'Xml')][string]$Method = 'Auto'
    )

    if (-not (Test-RunningOnWindows)) { throw 'Task Scheduler integration requires Windows.' }
    if (-not (Test-RunningAsAdministrator)) { throw 'Registering a scheduled task requires administrator privileges.' }
    if ([string]::IsNullOrWhiteSpace($ScriptPath) -or -not (Test-Path -LiteralPath $ScriptPath -PathType Leaf)) {
        throw "Entry script not found: $ScriptPath"
    }

    # Prefer the engine we are running under, so a PowerShell 7 deployment does
    # not silently fall back to 5.1 and lose -ThrottleLimit.
    $psExe = $null
    try { $psExe = (Get-Process -Id $PID -ErrorAction Stop).Path }
    catch { $psExe = $null }
    if ([string]::IsNullOrWhiteSpace($psExe) -or -not (Test-Path -LiteralPath $psExe)) {
        $psExe = Join-Path $PSHOME 'powershell.exe'
    }
    if (-not (Test-Path -LiteralPath $psExe)) { $psExe = 'powershell.exe' }

    $workDir = Split-Path -Parent $ScriptPath
    $taskArgs = New-SpoolerTaskArgumentString -ScriptPath $ScriptPath -ComputerName $ComputerName `
        -ComputerListFile $ComputerListFile -ExtraArguments $Arguments -Backup:$Backup -AutoUpdate:$AutoUpdate
    $limitMinutes = [Math]::Max(5, $ExecutionTimeLimitMinutes)
    $limit = 'PT{0}M' -f $limitMinutes

    $xml = New-SpoolerTaskXml -Command $psExe -Arguments $taskArgs -WorkingDirectory $workDir `
        -RunAsAccount $RunAsAccount -Schedule $Schedule -StartTime $StartTime -IntervalHours $IntervalHours `
        -EventLogName $EventLogName -EventSource $EventSource -EventId $EventId -ExecutionTimeLimit $limit `
        -RunOnlyIfIdle:$RunOnlyIfIdle `
        -Description ("Automated print spooler queue maintenance (v{0}). Clears stuck jobs, backs them up first, and restarts the Spooler safely." -f $Script:ScriptVersion)

    $registered = $false
    $usedMethod = $null
    $parts = Get-SpoolerTaskFolderAndLeaf -TaskName $TaskName

    if ($Method -in @('Auto', 'Module') -and (Get-Command -Name Register-ScheduledTask -ErrorAction SilentlyContinue)) {
        $action = $null; $trigger = $null; $principal = $null; $settings = $null
        try {
            $action = New-ScheduledTaskAction -Execute $psExe -Argument $taskArgs -WorkingDirectory $workDir

            if ($Schedule -eq 'OnEvent') {
                # There is no New-ScheduledTaskTrigger equivalent for an event
                # subscription, so build the CIM instance directly.
                $cimClass = Get-CimClass -ClassName MSFT_TaskEventTrigger -Namespace 'root/Microsoft/Windows/TaskScheduler' -ErrorAction Stop
                $trigger = New-CimInstance -CimClass $cimClass -ClientOnly
                $logEsc = [System.Security.SecurityElement]::Escape($EventLogName)
                $srcEsc = [System.Security.SecurityElement]::Escape($EventSource)
                $idPart = if ($EventId) { " and EventID=$EventId" } else { '' }
                $trigger.Subscription = "<QueryList><Query Id=""0"" Path=""$logEsc""><Select Path=""$logEsc"">*[System[Provider[@Name='$srcEsc']$idPart]]</Select></Query></QueryList>"
                $trigger.Enabled = $true
            }
            else {
                $trigger = switch ($Schedule) {
                    'Daily' { New-ScheduledTaskTrigger -Daily -At $StartTime }
                    'Weekly' { New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday -At $StartTime }
                    'AtStartup' { New-ScheduledTaskTrigger -AtStartup }
                    'Once' { New-ScheduledTaskTrigger -Once -At $StartTime }
                    'Hourly' { New-ScheduledTaskTrigger -Once -At $StartTime -RepetitionInterval (New-TimeSpan -Hours ([Math]::Max(1, $IntervalHours))) }
                    default { New-ScheduledTaskTrigger -Daily -At $StartTime }
                }
            }

            $principal = New-ScheduledTaskPrincipal -UserId $RunAsAccount -LogonType ServiceAccount -RunLevel Highest

            $settingsParams = @{
                AllowStartIfOnBatteries    = $true
                DontStopIfGoingOnBatteries = $true
                StartWhenAvailable         = $true
                ExecutionTimeLimit         = (New-TimeSpan -Minutes $limitMinutes)
                RestartCount               = 2
                RestartInterval            = (New-TimeSpan -Minutes 10)
                MultipleInstances          = 'IgnoreNew'
            }
            if ($RunOnlyIfIdle) { $settingsParams['RunOnlyIfIdle'] = $true }
            $settings = New-ScheduledTaskSettingsSet @settingsParams

            if ($parts.Folder -ne '\') { $null = New-SpoolerTaskFolder -Folder $parts.Folder }

            Register-ScheduledTask -TaskPath $parts.Folder -TaskName $parts.Leaf -Action $action -Trigger $trigger `
                -Principal $principal -Settings $settings -Force -ErrorAction Stop | Out-Null
            $registered = $true
            $usedMethod = 'Module'
        }
        catch {
            Write-LogWarn "Register-ScheduledTask failed ($($_.Exception.Message)); falling back to schtasks /xml." -Component 'Task'
        }
    }

    if (-not $registered) {
        $tmp = Join-Path ([IO.Path]::GetTempPath()) ("psq-task-{0}.xml" -f [Guid]::NewGuid().ToString('N'))
        try {
            # schtasks.exe requires UTF-16LE for /xml.
            [IO.File]::WriteAllText($tmp, $xml, [System.Text.Encoding]::Unicode)
            $proc = Start-Process -FilePath 'schtasks.exe' -ArgumentList @('/Create', '/TN', $TaskName, '/XML', $tmp, '/F') `
                -NoNewWindow -Wait -PassThru -ErrorAction Stop
            if ($proc.ExitCode -ne 0) { throw "schtasks.exe exited with code $($proc.ExitCode)." }
            $registered = $true
            $usedMethod = 'Xml'
        }
        catch {
            throw "Could not register the scheduled task: $($_.Exception.Message)"
        }
        finally {
            Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        }
    }

    Write-LogSuccess ("Scheduled task '{0}' registered via {1}." -f $TaskName, $usedMethod) -Component 'Task'

    if ($PassThru) {
        return [pscustomobject]@{
            TaskName           = $TaskName
            Method             = $usedMethod
            Command            = $psExe
            Arguments          = $taskArgs
            Schedule           = $Schedule
            StartTime          = $StartTime
            RunAsAccount       = $RunAsAccount
            ExecutionTimeLimit = $limit
            Xml                = $xml
            Registered         = $registered
            Task               = (Get-SpoolerCleanupTaskInternal -TaskName $TaskName)
        }
    }
    return $registered
}

function Get-SpoolerCleanupTaskInternal {
    <#
    .SYNOPSIS
        Reads back the registered task, with a schtasks /query fallback.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$TaskName)

    if (Get-Command -Name Get-ScheduledTask -ErrorAction SilentlyContinue) {
        $parts = Get-SpoolerTaskFolderAndLeaf -TaskName $TaskName
        try {
            $t = Get-ScheduledTask -TaskPath $parts.Folder -TaskName $parts.Leaf -ErrorAction Stop
            $i = $t | Get-ScheduledTaskInfo -ErrorAction SilentlyContinue
            return [pscustomobject]@{
                TaskName           = $TaskName
                Exists             = $true
                State              = [string]$t.State
                LastRunTime        = $i.LastRunTime
                LastTaskResult     = $i.LastTaskResult
                NextRunTime        = $i.NextRunTime
                NumberOfMissedRuns = $i.NumberOfMissedRuns
                Action             = ((@($t.Actions) | ForEach-Object { "$($_.Execute) $($_.Arguments)" }) -join ' ; ')
            }
        }
        catch {
            return [pscustomobject]@{ TaskName = $TaskName; Exists = $false; State = 'NotFound'; Error = $_.Exception.Message }
        }
    }

    try {
        $out = & schtasks.exe /Query /TN $TaskName /FO LIST /V 2>&1
        if ($LASTEXITCODE -ne 0) { return [pscustomobject]@{ TaskName = $TaskName; Exists = $false; State = 'NotFound' } }
        return [pscustomobject]@{ TaskName = $TaskName; Exists = $true; State = 'Registered'; Raw = ($out -join [Environment]::NewLine) }
    }
    catch {
        return [pscustomobject]@{ TaskName = $TaskName; Exists = $false; State = 'Unknown'; Error = $_.Exception.Message }
    }
}

function Uninstall-SpoolerCleanupTaskInternal {
    <#
    .SYNOPSIS
        Removes the registered task.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$TaskName)

    if (-not (Test-RunningOnWindows)) { throw 'Task Scheduler integration requires Windows.' }

    if (Get-Command -Name Unregister-ScheduledTask -ErrorAction SilentlyContinue) {
        try {
            $parts = Get-SpoolerTaskFolderAndLeaf -TaskName $TaskName
            Unregister-ScheduledTask -TaskPath $parts.Folder -TaskName $parts.Leaf -Confirm:$false -ErrorAction Stop
            return $true
        }
        catch {
            Write-LogWarn "Unregister-ScheduledTask failed ($($_.Exception.Message)); trying schtasks /delete." -Component 'Task'
        }
    }

    try {
        $null = & schtasks.exe /Delete /TN $TaskName /F 2>&1
        return ($LASTEXITCODE -eq 0)
    }
    catch {
        return $false
    }
}

function Start-SpoolerCleanupTaskInternal {
    <#
    .SYNOPSIS
        Starts the registered task and optionally waits for its result.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$TaskName,
        [switch]$Wait,
        [int]$TimeoutSeconds = 600
    )

    $out = [ordered]@{ Started = $false; TaskName = $TaskName; ExitCode = $null; Waited = $false }
    try {
        if (Get-Command -Name Start-ScheduledTask -ErrorAction SilentlyContinue) {
            $parts = Get-SpoolerTaskFolderAndLeaf -TaskName $TaskName
            Start-ScheduledTask -TaskPath $parts.Folder -TaskName $parts.Leaf -ErrorAction Stop
        }
        else {
            $null = & schtasks.exe /Run /TN $TaskName 2>&1
            if ($LASTEXITCODE -ne 0) { throw "schtasks /Run failed with code $LASTEXITCODE." }
        }
        $out.Started = $true
        Write-LogInfo "Scheduled task '$TaskName' started." -Component 'Task'

        if ($Wait) {
            $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
            while ((Get-Date) -lt $deadline) {
                $info = Get-SpoolerCleanupTaskInternal -TaskName $TaskName
                if ($info.State -ne 'Running') {
                    $out.ExitCode = $info.LastTaskResult
                    $out.Waited = $true
                    break
                }
                Start-Sleep -Seconds 2
            }
            if (-not $out.Waited) { Write-LogWarn "Timed out after ${TimeoutSeconds}s waiting for '$TaskName'." -Component 'Task' }
        }
    }
    catch {
        Write-LogError "Could not start '$TaskName': $($_.Exception.Message)" -Component 'Task'
    }
    return [pscustomobject]$out
}

# ============================================================================
#  REMOTE/LOCAL EXECUTION PAYLOAD  (template - token-substituted, never run here)
# ----------------------------------------------------------------------------
#  WHY ONE PAYLOAD FOR BOTH CASES
#  v3 kept two divergent code paths (local cmdlets vs. Invoke-Command), which is
#  exactly how the v3.0 bug "Clear-SpoolFiles referenced an undefined
#  $ScriptConfig.SpoolDirectory" survived review. v4 has ONE implementation that
#  runs on the target machine; the only difference is whether it is invoked
#  directly (local) or through Invoke-Command (remote). Behaviour parity is
#  therefore structural, not aspirational.
#
#  CONTRACT
#  In : $Options - a flat hashtable of scalars/arrays (CLIXML-serializable).
#  Out: a single hashtable with the keys documented in
#       docs/Report-Schema.md. Always exactly one object on the output stream.
#
#  The payload must only use cmdlets available in Windows PowerShell 5.1 so the
#  same code runs on an old Server 2012 R2 target and on PowerShell 7.6.
# ============================================================================

$Script:RemotePayloadTemplate = @'

    # The dispatcher invokes this block as `& $Payload $Options`, so the options
    # hashtable arrives as the FIRST argument. Without the declaration below the value
    # lands in $args and every $Options read fails - and the block must NOT be wrapped
    # in an extra pair of braces, or invoking it merely returns the inner scriptblock
    # and the cleanup never runs at all.
    param($Options)
    # ---- payload scope ------------------------------------------------------
    $ErrorActionPreference = 'Continue'
    Set-StrictMode -Version 2.0

    $sw        = [System.Diagnostics.Stopwatch]::StartNew()
    $spoolSvc  = '@@SPOOLER_SERVICE@@'
    $depNames  = @(@@DEPENDENT_SERVICES@@)
    $exts      = @(@@SPOOL_EXTENSIONS@@) | ForEach-Object { $_.ToUpperInvariant() }
    $legacyExt = @(@@LEGACY_EXTENSIONS@@) | ForEach-Object { $_.ToUpperInvariant() }

    $R = @{
        ComputerName         = $env:COMPUTERNAME
        Status               = 'Unknown'
        StartedUtc           = [DateTime]::UtcNow.ToString('o')
        FinishedUtc          = $null
        DurationMs           = 0
        SpoolDirectory       = $null
        SpoolDirectorySource = $null
        SpoolDirectoryValid  = $false
        ServiceWasRunning    = $null
        ServiceIsRunning     = $null
        ServiceStopMs        = 0
        ServiceStartMs       = 0
        DependentServices    = @()
        JobsFound            = 0
        JobsSelected         = 0
        JobsRemoved          = 0
        FilesFound           = 0
        FilesSelected        = 0
        FilesDeleted         = 0
        FilesFailed          = 0
        FilesLocked          = 0
        BytesFound           = [long]0
        BytesDeleted         = [long]0
        FilesAfter           = 0
        ResurrectedFiles     = 0
        BackupPath           = $null
        BackupHash           = $null
        DryRun               = [bool]$Options['DryRun']
        RepairPerformed      = $false
        DiagnosticsCollected = $false
        Mode                 = $Options['Mode']
        Warnings             = New-Object System.Collections.Generic.List[string]
        Errors               = New-Object System.Collections.Generic.List[string]
        Jobs                 = New-Object System.Collections.Generic.List[object]
        RemovedFiles         = New-Object System.Collections.Generic.List[object]
        FailedFiles          = New-Object System.Collections.Generic.List[object]
        OrphanedFiles        = New-Object System.Collections.Generic.List[string]
        Diagnostics          = $null
        ServiceTrace         = New-Object System.Collections.Generic.List[object]
    }

    # Every return path funnels through here so the emitted record always has
    # exactly the same shape (Lists are converted to arrays: CLIXML round-trips
    # arrays reliably, generic Lists do not always deserialize as expected).
    function Finish-Result([string]$status) {
        $R['Status']       = $status
        $R['Warnings']     = @($R['Warnings'].ToArray())
        $R['Errors']       = @($R['Errors'].ToArray())
        $R['Jobs']         = @($R['Jobs'])
        $R['RemovedFiles'] = @($R['RemovedFiles'].ToArray())
        $R['FailedFiles']  = @($R['FailedFiles'].ToArray())
        $R['OrphanedFiles']= @($R['OrphanedFiles'].ToArray())
        $R['ServiceTrace'] = @($R['ServiceTrace'].ToArray())
        $sw.Stop()
        $R['DurationMs']   = $sw.ElapsedMilliseconds
        $R['FinishedUtc']  = [DateTime]::UtcNow.ToString('o')
        return $R
    }

    function Add-W([string]$m) { $R['Warnings'].Add($m) | Out-Null }
    function Add-E([string]$m) { $R['Errors'].Add($m)   | Out-Null }
    function Trace([string]$svc, [string]$action, [string]$from, [string]$to, [bool]$ok, [string]$note) {
        $R['ServiceTrace'].Add([pscustomobject]@{
            Service = $svc; Action = $action; From = $from; To = $to; Success = $ok
            TimeUtc = [DateTime]::UtcNow.ToString('o'); Note = $note
        }) | Out-Null
    }

    # ------------------------------------------------------------------------
    #  Service state helpers
    # ------------------------------------------------------------------------
    function Get-SvcState([string]$name) {
        # Win32_Service first: it reports StartMode and the owning PID, which the
        # .NET ServiceController does not. Get-Service is the fallback for hosts
        # where the CIM provider is unavailable.
        try {
            $c = Get-CimInstance -ClassName Win32_Service -Filter "Name='$name'" -ErrorAction Stop
            if ($c) {
                return [pscustomobject]@{ Exists = $true; State = [string]$c.State; StartMode = [string]$c.StartMode; ProcessId = [int]$c.ProcessId }
            }
        }
        catch {
            # A missing CIM provider is not an error; fall through to Get-Service.
        }
        try {
            $s = Get-Service -Name $name -ErrorAction Stop
            return [pscustomobject]@{ Exists = $true; State = [string]$s.Status; StartMode = [string]$s.StartType; ProcessId = 0 }
        }
        catch {
            # Not installed, or no access: report "does not exist" rather than
            # throwing, so the caller can treat it as a no-op.
        }
        return [pscustomobject]@{ Exists = $false; State = 'Unknown'; StartMode = 'Unknown'; ProcessId = 0 }
    }

    function Resolve-DependentServiceNames {
        $found = New-Object System.Collections.Generic.List[string]
        foreach ($n in $depNames) {
            if ($n -eq 'PrintWorkflowUserSvc') {
                foreach ($s in @(Get-Service -ErrorAction SilentlyContinue | Where-Object { $_.Name -like 'PrintWorkflowUserSvc*' })) {
                    $found.Add($s.Name) | Out-Null
                }
            } else {
                $st = Get-SvcState $n
                if ($st.Exists) { $found.Add($n) | Out-Null }
            }
        }
        return $found.ToArray()
    }

    function Stop-OneService([string]$name, [int]$timeoutSec, [switch]$AllowKill) {
        $before = Get-SvcState $name
        if (-not $before.Exists) { Trace $name 'Stop' 'Missing' 'Missing' $true 'Service not installed'; return @{ Ok = $true; WasRunning = $false; State = 'Missing' } }
        if ($before.State -eq 'Stopped') { Trace $name 'Stop' 'Stopped' 'Stopped' $true 'Already stopped'; return @{ Ok = $true; WasRunning = $false; State = 'Stopped' } }

        $stopSw = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            Stop-Service -Name $name -Force -ErrorAction Stop
        }
        catch {
            # Stop-Service throws on a pending/dependent-service problem, but the
            # poll below is the real arbiter of success, so record and continue.
            Add-W ("Stop-Service '{0}' threw: {1}" -f $name, $_.Exception.Message)
        }

        $deadline = (Get-Date).AddSeconds($timeoutSec)
        $state = $before.State
        while ((Get-Date) -lt $deadline) {
            $state = (Get-SvcState $name).State
            if ($state -eq 'Stopped') { break }
            Start-Sleep -Milliseconds 250
        }

        if (($state -ne 'Stopped') -and $AllowKill) {
            # FINDING A1/A4: PrintNotify and a hung splwow64.exe keep the Spooler in
            # STOP_PENDING forever. Escalate to a process kill, then re-poll. The
            # SCM treats an exited service process as stopped, so this is safe.
            $pid0 = $before.ProcessId
            if ($pid0 -gt 0) {
                try {
                    Stop-Process -Id $pid0 -Force -ErrorAction Stop
                    Add-W ("Service '{0}' did not stop within {1}s; terminated PID {2}." -f $name, $timeoutSec, $pid0)
                }
                catch {
                    Add-W ("Could not terminate PID {0} for '{1}': {2}" -f $pid0, $name, $_.Exception.Message)
                }
                $deadline2 = (Get-Date).AddSeconds([Math]::Max(5, [int]($timeoutSec / 2)))
                while ((Get-Date) -lt $deadline2) {
                    $state = (Get-SvcState $name).State
                    if ($state -eq 'Stopped') { break }
                    Start-Sleep -Milliseconds 250
                }
            }
        }
        $stopSw.Stop()

        $ok = ($state -eq 'Stopped')
        Trace $name 'Stop' $before.State $state $ok ("{0} ms" -f $stopSw.ElapsedMilliseconds)
        return @{ Ok = $ok; WasRunning = $true; State = $state; ElapsedMs = $stopSw.ElapsedMilliseconds }
    }

    function Start-OneService([string]$name, [int]$timeoutSec) {
        $before = Get-SvcState $name
        if (-not $before.Exists) { Trace $name 'Start' 'Missing' 'Missing' $true 'Service not installed'; return $true }
        if ($before.State -eq 'Running') { Trace $name 'Start' 'Running' 'Running' $true 'Already running'; return $true }

        $startSw = [System.Diagnostics.Stopwatch]::StartNew()
        try { Start-Service -Name $name -ErrorAction Stop } catch { Add-W ("Start-Service '{0}' threw: {1}" -f $name, $_.Exception.Message) }
        $deadline = (Get-Date).AddSeconds($timeoutSec)
        $state = $before.State
        while ((Get-Date) -lt $deadline) {
            $state = (Get-SvcState $name).State
            if ($state -eq 'Running') { break }
            Start-Sleep -Milliseconds 250
        }
        $startSw.Stop()
        $ok = ($state -eq 'Running')
        Trace $name 'Start' $before.State $state $ok ("{0} ms" -f $startSw.ElapsedMilliseconds)
        return $ok
    }

    # ------------------------------------------------------------------------
    #  Spool directory resolution + SAFETY GUARDS
    # ------------------------------------------------------------------------
    function Resolve-SpoolDir {
        if ($Options['SpoolDirectoryOverride']) {
            return @{ Path = [string]$Options['SpoolDirectoryOverride']; Source = 'Parameter'; RegistryValue = $null }
        }
        $def = Join-Path $env:SystemRoot '@@DEFAULT_SPOOL_SUBPATH@@'
        $regVal = $null
        try {
            $key = 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers'
            if (Test-Path $key) {
                $item = Get-ItemProperty -Path $key -Name 'DefaultSpoolDirectory' -ErrorAction SilentlyContinue
                if ($item -and $item.DefaultSpoolDirectory) { $regVal = [string]$item.DefaultSpoolDirectory }
            }
        } catch { Add-W ("Registry probe failed: {0}" -f $_.Exception.Message) }

        if ($regVal -and $regVal.Trim()) {
            $t = $regVal.Trim().Trim('"')
            if ($t -match '^\\\\' -or $t.Length -lt 4 -or $t -match '[\x00-\x1f]') {
                Add-W ("Ignoring suspicious DefaultSpoolDirectory '{0}'; using OS default." -f $regVal)
            } else {
                return @{ Path = $t; Source = 'Registry'; RegistryValue = $regVal }
            }
        }
        return @{ Path = $def; Source = 'OSDefault'; RegistryValue = $regVal }
    }

    function Test-SpoolDirSafe([string]$dir) {
        # Four independent guards. ANY failure aborts before a single delete.
        if ([string]::IsNullOrWhiteSpace($dir)) { return 'Spool directory path is empty.' }
        if ($dir -match '[\x00-\x1f]')          { return 'Spool directory path contains control characters.' }
        if ($dir -match '^\\\\')                { return 'UNC spool directories are refused by policy (remote file systems must be cleaned on their owner).' }
        if ($dir -notmatch '^[A-Za-z]:[\\\/]')     { return 'Spool directory is not rooted at a Windows drive letter.' }
        if ($dir.Substring(2) -match '[<>:"|?*]')   { return 'Spool directory contains characters that are invalid in a Windows path.' }

        $full = $null
        try { $full = [System.IO.Path]::GetFullPath($dir) } catch { return "Spool directory path cannot be resolved: $dir" }
        if (-not [System.IO.Path]::IsPathRooted($full)) { return "Spool directory path is not rooted: $full" }

        $leaf = Split-Path -Leaf $full.TrimEnd('\', '/')
        if ($leaf -notlike '@@GUARD_REQUIRED_LEAF@@') {
            return "Refusing to operate on '$full': the resolved spool directory must end with '@@GUARD_REQUIRED_LEAF@@' (got '$leaf'). Override deliberately with -SpoolDirectory only if this host uses a non-standard layout."
        }

        $depth = @($dir.Trim().TrimEnd('\', '/').Split([char[]]@('\', '/')) | Where-Object { $_ }).Count
        if ($depth -lt @@GUARD_MIN_DEPTH@@) { return "Refusing to operate on '$full': path is too shallow ($depth segments)." }

        if ($full -notmatch '(?i)[\\\/]spool[\\\/]') {
            return "Refusing to operate on '$full': a spool directory must live under a 'spool' folder."
        }

        foreach ($bad in @(@@FORBIDDEN_ROOTS@@)) {
            if (-not $bad) { continue }
            $bn = $bad.TrimEnd('\', '/')
            if ([string]::Equals($full.TrimEnd('\', '/'), $bn, [StringComparison]::OrdinalIgnoreCase)) {
                return "Refusing to operate on '$full': the path equals a protected system root."
            }
        }
        return $null
    }

    # ------------------------------------------------------------------------
    #  Job inventory
    # ------------------------------------------------------------------------
    function Get-PrintJobsWmi {
        # Authoritative source, but only answers while the Spooler is alive.
        $list = New-Object System.Collections.Generic.List[object]
        try {
            $jobs = @(Get-CimInstance -ClassName Win32_PrintJob -ErrorAction Stop)
            foreach ($j in $jobs) {
                $name = [string]$j.Name           # e.g. "HP LaserJet, 12"
                $jid  = 0
                $pname = $name
                if ($name -match '^(?<p>.+),\s*(?<i>\d+)$') { $pname = $Matches['p']; $jid = [int]$Matches['i'] }
                $submitted = $null
                try { if ($j.TimeSubmitted) { $submitted = [Management.ManagementDateTimeConverter]::ToDateTime([string]$j.TimeSubmitted) } } catch { /* an unparseable timestamp leaves the field null rather than failing */ }
                $list.Add([pscustomobject]@{
                    JobId        = $jid
                    PrinterName  = $pname
                    DocumentName = [string]$j.Document
                    UserName     = [string]$j.Owner
                    SizeBytes    = [long]($j.Size)
                    TotalPages   = [int]($j.TotalPages)
                    PagesPrinted = [int]($j.PagesPrinted)
                    Status       = [string]$j.JobStatus
                    Priority     = [int]$j.Priority
                    SubmittedUtc = $submitted
                    Source       = 'WMI'
                    ShdFile      = $null
                    SplFile      = $null
                }) | Out-Null
            }
        } catch {
            Add-W ("WMI Win32_PrintJob enumeration failed (spooler hung?): {0}" -f $_.Exception.Message)
        }
        return $list.ToArray()
    }

    function Get-ShdPrintableStrings {
        <#
            Best-effort extraction of the human-readable fields embedded in a
            .SHD shadow file. The SHD header layout changed across Windows
            versions, so instead of marshalling a version-specific struct we
            scan for the two shapes that are stable in practice:
              * the UNC printer path  \\<machine>\<printer>
              * the source document path  <drive>:\<...>\<file.ext>
            Both are stored as UTF-16LE (Unicode SHD) or ANSI (legacy SHD).
            Accuracy is documented in the wiki; WMI remains the primary source.
        #>
        param([string]$Path)
        $out = @{ PrinterName = $null; DocumentName = $null; UserName = $null }
        try {
            $fi = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
            $len = [Math]::Min($fi.Length, 262144)
            if ($len -le 0) { return $out }
            $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            try {
                $buf = New-Object byte[] $len
                $read = $fs.Read($buf, 0, $len)
                if ($read -le 0) { return $out }
                $uni  = [System.Text.Encoding]::Unicode.GetString($buf, 0, $read)
                $ansi = [System.Text.Encoding]::Default.GetString($buf, 0, $read)
                foreach ($text in @($uni, $ansi)) {
                    if (-not $out.PrinterName -and $text -match '\\\\(?<m>[^\\\x00-\x1f]{1,64})\\(?<p>[^\\\x00-\x1f]{1,128})') {
                        $out.PrinterName = $Matches['p']
                    }
                    if (-not $out.DocumentName -and $text -match '(?<d>[A-Za-z]:\\[^\x00-\x1f\\]{1,200}(?:\\[^\x00-\x1f\\]{1,200})*\.[A-Za-z0-9]{1,10})') {
                        $out.DocumentName = $Matches['d']
                    }
                    if ($out.PrinterName -and $out.DocumentName) { break }
                }
            } finally { $fs.Dispose() }
        }
        catch {
            /* an unreadable or exclusively-locked .SHD simply yields no metadata;
               the caller falls back to the file-name-derived job id */
        }
        return $out
    }

    function Get-PrintJobsFromSpoolDir {
        <#
            Works even when the Spooler is completely hung (the normal case for
            the tickets this tool exists to close). Pairs F#####.SHD with
            the matching .SPL and derives the job id from the file name.
        #>
        param([string]$Dir)
        $list = New-Object System.Collections.Generic.List[object]
        try {
            $shd = @(Get-ChildItem -LiteralPath $Dir -File -Force -ErrorAction SilentlyContinue | Where-Object { $_.Extension -ieq '.SHD' })
            foreach ($f in $shd) {
                $base = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
                $jid = 0
                if ($base -match '(\d+)$') { $jid = [int64]$Matches[1] }
                $meta = Get-ShdPrintableStrings -Path $f.FullName
                $spl = $null
                foreach ($e in @('.SPL', '.FPF')) {
                    $cand = Join-Path $Dir ($base + $e)
                    if (Test-Path -LiteralPath $cand) { $spl = $cand; break }
                }
                $size = [long]$f.Length
                if ($spl) {
                    try { $size += (Get-Item -LiteralPath $spl -Force).Length }
                    catch { /* the .SPL size is a nice-to-have in the inventory */ }
                }
                $list.Add([pscustomobject]@{
                    JobId        = $jid
                    PrinterName  = $meta.PrinterName
                    DocumentName = $meta.DocumentName
                    UserName     = $meta.UserName
                    SizeBytes    = $size
                    TotalPages   = 0
                    PagesPrinted = 0
                    Status       = 'Unknown'
                    Priority     = 0
                    SubmittedUtc = $f.LastWriteTimeUtc
                    Source       = 'SHD'
                    ShdFile      = $f.Name
                    SplFile      = $(if ($spl) { [System.IO.Path]::GetFileName($spl) } else { $null })
                }) | Out-Null
            }
        } catch { Add-E ("Spool directory inventory failed: {0}" -f $_.Exception.Message) }
        return $list.ToArray()
    }

    function Select-JobsToActOn {
        <#
            Applies the operator's selection filters. Returns the subset plus a
            note explaining WHY the selection may be unreliable (this is the
            difference between a tool and a hazard).
        #>
        param($Jobs)
        $sel = @($Jobs)
        $note = $null

        if ($Options['PrinterName'] -and @($Options['PrinterName']).Count -gt 0) {
            $pats = @($Options['PrinterName'])
            $sel = @($sel | Where-Object {
                $pn = [string]$_.PrinterName
                if ([string]::IsNullOrWhiteSpace($pn)) { return $false }
                foreach ($p in $pats) { if ($pn -like $p) { return $true } }
                return $false
            })
        }
        if ($Options['ExcludePrinterName'] -and @($Options['ExcludePrinterName']).Count -gt 0) {
            $xpats = @($Options['ExcludePrinterName'])
            $sel = @($sel | Where-Object {
                $pn = [string]$_.PrinterName
                if ([string]::IsNullOrWhiteSpace($pn)) { return $true }
                foreach ($p in $xpats) { if ($pn -like $p) { return $false } }
                return $true
            })
        }
        if ($Options['JobId'] -and @($Options['JobId']).Count -gt 0) {
            $ids = @($Options['JobId'] | ForEach-Object { [long]$_ })
            $sel = @($sel | Where-Object { $ids -contains [long]$_.JobId })
        }
        if ($Options['DocumentName']) {
            $dp = [string]$Options['DocumentName']
            $sel = @($sel | Where-Object { ([string]$_.DocumentName) -like $dp })
        }
        if ($Options['OwnerName']) {
            $op = [string]$Options['OwnerName']
            $sel = @($sel | Where-Object { ([string]$_.UserName) -like $op })
        }
        if ($Options['OlderThanSeconds'] -gt 0) {
            $cut = [DateTime]::UtcNow.AddSeconds(-[double]$Options['OlderThanSeconds'])
            $sel = @($sel | Where-Object {
                if (-not $_.SubmittedUtc) { return $false }
                ([DateTime]$_.SubmittedUtc).ToUniversalTime() -lt $cut
            })
        }
        if ($Options['NewerThanSeconds'] -gt 0) {
            $cut = [DateTime]::UtcNow.AddSeconds(-[double]$Options['NewerThanSeconds'])
            $sel = @($sel | Where-Object {
                if (-not $_.SubmittedUtc) { return $false }
                ([DateTime]$_.SubmittedUtc).ToUniversalTime() -gt $cut
            })
        }
        if ($Options['LargerThanBytes'] -gt 0) {
            $lim = [long]$Options['LargerThanBytes']
            $sel = @($sel | Where-Object { [long]$_.SizeBytes -gt $lim })
        }
        if ($Options['SmallerThanBytes'] -gt 0) {
            $lim = [long]$Options['SmallerThanBytes']
            $sel = @($sel | Where-Object { [long]$_.SizeBytes -lt $lim })
        }

        $selective = ($Options['PrinterName'] -or $Options['ExcludePrinterName'] -or $Options['JobId'] -or
                      $Options['DocumentName'] -or $Options['OwnerName'] -or $Options['OlderThanSeconds'] -gt 0 -or
                      $Options['NewerThanSeconds'] -gt 0 -or $Options['LargerThanBytes'] -gt 0 -or $Options['SmallerThanBytes'] -gt 0)
        if ($selective) {
            $unknownMeta = @($Jobs | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.PrinterName) }).Count
            if ($unknownMeta -gt 0) {
                $note = ("{0} of {1} queued jobs have no readable printer/document metadata, so they cannot be matched by name-based filters and were NOT selected. Use -All when the whole queue must go." -f $unknownMeta, @($Jobs).Count)
            }
        }
        return @{ Selected = $sel; Note = $note; Selective = $selective }
    }

    function Get-FilesForJob {
        param($Job, [string]$Dir)
        $files = New-Object System.Collections.Generic.List[string]
        if ($Job.ShdFile) { $p = Join-Path $Dir $Job.ShdFile; if (Test-Path -LiteralPath $p) { $files.Add($p) | Out-Null } }
        if ($Job.SplFile) { $p = Join-Path $Dir $Job.SplFile; if (Test-Path -LiteralPath $p) { $files.Add($p) | Out-Null } }
        if ($files.Count -eq 0 -and $Job.JobId -gt 0) {
            foreach ($e in $exts) {
                $cand = Join-Path $Dir ('{0}{1:D5}{2}' -f $(if ($e -eq '.SHD') { 'F' } else { '' }), [int]$Job.JobId, $e)
                if (Test-Path -LiteralPath $cand) { $files.Add($cand) | Out-Null }
            }
        }
        return $files.ToArray()
    }

    # ------------------------------------------------------------------------
    #  Deletion with retry/backoff and lock classification
    # ------------------------------------------------------------------------
    function Remove-SpoolFileRobust {
        param([string]$Path, [int]$RetryCount, [int]$RetryDelayMs)
        $attempt = 0
        $lastErr = $null
        while ($true) {
            try {
                if (-not (Test-Path -LiteralPath $Path)) { return @{ Ok = $true; Locked = $false; Attempts = $attempt; Error = $null } }
                $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
                if ($item.Attributes -band [IO.FileAttributes]::ReadOnly) {
                    try { Set-ItemProperty -LiteralPath $Path -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue } catch { /* clearing the read-only attribute is a best effort before the delete retry */ }
                }
                [System.IO.File]::Delete($Path)
                if (-not (Test-Path -LiteralPath $Path)) { return @{ Ok = $true; Locked = $false; Attempts = ($attempt + 1); Error = $null } }
                $lastErr = 'File still present after delete'
            } catch {
                $lastErr = $_.Exception.Message
                $hresult = $null
                try { $hresult = $_.Exception.HResult } catch { /* HResult is not exposed on every exception type */ }
                # 0x80070020 sharing violation / 0x80070005 access denied => locked
            }
            $attempt++
            if ($attempt -gt $RetryCount) { break }
            $delay = [int]($RetryDelayMs * [Math]::Pow(@@RETRY_BACKOFF@@, $attempt - 1))
            Start-Sleep -Milliseconds ([Math]::Min(10000, $delay))
        }
        $locked = ($lastErr -match 'being used by another process|sharing violation|Access to the path|denied|Used by another')
        return @{ Ok = $false; Locked = $locked; Attempts = $attempt; Error = $lastErr }
    }

    # ------------------------------------------------------------------------
    #  Backup (finding A3)
    # ------------------------------------------------------------------------
    function New-SpoolBackup {
        param($Files, [string]$BackupRoot, [string]$Computer, $SelectedJobs)
        $result = @{ Path = $null; Hash = $null; FilesCopied = 0; Error = $null; Manifest = $null }
        try {
            if (-not $BackupRoot) { $BackupRoot = Join-Path $env:ProgramData '@@PRODUCT_FAMILY@@\Backups' }
            if (-not (Test-Path -LiteralPath $BackupRoot)) { New-Item -ItemType Directory -Path $BackupRoot -Force -ErrorAction Stop | Out-Null }

            $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
            $safeComputer = ($Computer -replace '[^A-Za-z0-9._-]', '_')
            $bundleName = ('spool-backup_{0}_{1}' -f $safeComputer, $stamp)
            $staging = Join-Path ([System.IO.Path]::GetTempPath()) $bundleName
            if (Test-Path -LiteralPath $staging) { Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue }
            New-Item -ItemType Directory -Path $staging -Force -ErrorAction Stop | Out-Null

            $copied = 0
            foreach ($f in $Files) {
                try { Copy-Item -LiteralPath $f -Destination (Join-Path $staging ([System.IO.Path]::GetFileName($f))) -Force -ErrorAction Stop; $copied++ }
                catch { Add-W ("Backup: could not copy '{0}': {1}" -f $f, $_.Exception.Message) }
            }

            $manifest = [ordered]@{
                Tool            = '@@SCRIPT_NAME@@'
                ToolVersion     = '@@SCRIPT_VERSION@@'
                CreatedUtc      = [DateTime]::UtcNow.ToString('o')
                ComputerName    = $Computer
                SpoolDirectory  = $R['SpoolDirectory']
                FileCount       = $copied
                Files           = @($Files | ForEach-Object { @{ Name = [System.IO.Path]::GetFileName($_); Size = (Get-Item -LiteralPath $_ -Force -ErrorAction SilentlyContinue).Length } })
                Jobs            = @($SelectedJobs | ForEach-Object {
                    @{ JobId = $_.JobId; Printer = $_.PrinterName; Document = $_.DocumentName; Owner = $_.UserName; Size = $_.SizeBytes; Submitted = $_.SubmittedUtc }
                })
                RestoreNote     = 'Spool files cannot be re-queued by copying them back. Restore to a folder for forensic re-submission; see the wiki.'
            }
            $manifestPath = Join-Path $staging 'backup-manifest.json'
            $manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding UTF8 -ErrorAction Stop

            $zipPath = Join-Path $BackupRoot ($bundleName + '.zip')
            if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue }
            if (Get-Command -Name Compress-Archive -ErrorAction SilentlyContinue) {
                Compress-Archive -Path (Join-Path $staging '*') -DestinationPath $zipPath -CompressionLevel Optimal -ErrorAction Stop
            } else {
                Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
                [System.IO.Compression.ZipFile]::CreateFromDirectory($staging, $zipPath)
            }
            Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue

            $hash = $null
            try { $hash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256 -ErrorAction Stop).Hash } catch { /* a missing hash degrades the report, it must not abort the backup */ }

            $result.Path = $zipPath; $result.Hash = $hash; $result.FilesCopied = $copied
            $result.Manifest = $manifestPath
        } catch {
            $result.Error = $_.Exception.Message
        }
        return $result
    }

    # ------------------------------------------------------------------------
    #  Diagnostics payload (read-only)
    # ------------------------------------------------------------------------
    function Get-DiagBundle {
        param([string]$Dir, [int]$EventLogHours, [int]$EventLogMax)
        $d = [ordered]@{}
        try {
            $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
            $d['OS'] = @{ Caption = $os.Caption; Version = [string]$os.Version; Build = [string]$os.BuildNumber
                          LastBoot = if ($os.LastBootUpTime) { [Management.ManagementDateTimeConverter]::ToDateTime([string]$os.LastBootUpTime).ToString('o') } else { $null }
                          FreeMemoryMB = [int]($os.FreePhysicalMemory / 1024) }
        } catch { $d['OS'] = @{ Error = $_.Exception.Message } }

        $d['Services'] = @()
        foreach ($n in (@($spoolSvc) + @($depNames | Where-Object { $_ -ne 'PrintWorkflowUserSvc' }))) {
            $st = Get-SvcState $n
            $d['Services'] += [pscustomobject]@{ Name = $n; Exists = $st.Exists; State = $st.State; StartMode = $st.StartMode; ProcessId = $st.ProcessId }
        }
        foreach ($s in @(Get-Service -ErrorAction SilentlyContinue | Where-Object { $_.Name -like 'PrintWorkflowUserSvc*' })) {
            $d['Services'] += [pscustomobject]@{ Name = $s.Name; Exists = $true; State = [string]$s.Status; StartMode = [string]$s.StartType; ProcessId = 0 }
        }

        try {
            $d['SpoolerProcess'] = @(Get-Process -Name spoolsv, splwow64 -ErrorAction SilentlyContinue |
                Select-Object Name, Id, @{ n = 'CPUSeconds'; e = { [Math]::Round($_.CPU, 1) } },
                                      @{ n = 'WorkingSetMB'; e = { [Math]::Round($_.WorkingSet64 / 1MB, 1) } },
                                      @{ n = 'Threads'; e = { $_.Threads.Count } },
                                      @{ n = 'StartTime'; e = { try { $_.StartTime.ToString('o') } catch { $null } } },
                                      @{ n = 'Responding'; e = { $_.Responding } })
        } catch { $d['SpoolerProcess'] = @() }

        try {
            $d['Printers'] = @(Get-CimInstance Win32_Printer -ErrorAction SilentlyContinue | Select-Object Name, PrinterStatus, JobCountSinceLastReset,
                    Local, Network, Shared, PortName, DriverName, WorkOffline, PrinterState)
        } catch { $d['Printers'] = @() }

        try { $d['Drivers'] = @(Get-CimInstance Win32_PrinterDriver -ErrorAction SilentlyContinue | Select-Object Name, Manufacturer, DriverVersion, InfName) }
        catch { $d['Drivers'] = @() }

        try { $d['Ports'] = @(Get-CimInstance Win32_TCPIPPrinterPort -ErrorAction SilentlyContinue | Select-Object Name, HostAddress, PortNumber, Protocol, SNMPEnabled) }
        catch { $d['Ports'] = @() }

        if ($Dir -and (Test-Path -LiteralPath $Dir)) {
            try {
                $files = @(Get-ChildItem -LiteralPath $Dir -File -Force -ErrorAction SilentlyContinue)
                $oldest = $files | Sort-Object LastWriteTimeUtc | Select-Object -First 1
                $newest = $files | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
                # Summed explicitly. Measure-Object returns $null for an EMPTY
                # pipeline, and `.Sum` on $null throws under Set-StrictMode -Version
                # Latest - a hardening many production hosts enable. An empty spool
                # directory is the NORMAL case, not an edge case, so this must not
                # depend on lax mode.
                $dirBytes = [long]0
                foreach ($df in $files) { $dirBytes += [long]$df.Length }
                $d['SpoolDirectory'] = @{
                    Path = $Dir; FileCount = $files.Count
                    TotalBytes = $dirBytes
                    ByExtension = @($files | Group-Object { $_.Extension.ToUpperInvariant() } | ForEach-Object { $gb = [long]0; foreach ($gf in $_.Group) { $gb += [long]$gf.Length }; @{ Ext = $_.Name; Count = $_.Count; Bytes = $gb } })
                    OldestFile = if ($oldest) { @{ Name = $oldest.Name; AgeHours = [Math]::Round(((Get-Date) - $oldest.LastWriteTime).TotalHours, 1) } } else { $null }
                    NewestFile = if ($newest) { @{ Name = $newest.Name; AgeHours = [Math]::Round(((Get-Date) - $newest.LastWriteTime).TotalHours, 1) } } else { $null }
                    AclOwner = try { (Get-Acl -LiteralPath $Dir).Owner } catch { $null }
                }
            } catch { $d['SpoolDirectory'] = @{ Path = $Dir; Error = $_.Exception.Message } }
        }

        try {
            $reg = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers' -ErrorAction SilentlyContinue
            $d['Registry'] = @{
                DefaultSpoolDirectory = $reg.DefaultSpoolDirectory
                EventLog              = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Print' -Name 'EventLog' -ErrorAction SilentlyContinue).EventLog
                PrintDriversIsolation = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Printers' -ErrorAction SilentlyContinue)
            }
        } catch { $d['Registry'] = @{ Error = $_.Exception.Message } }

        if ($EventLogHours -gt 0) {
            $since = (Get-Date).AddHours(-$EventLogHours)
            try {
                $d['EventLog'] = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; StartTime = $since; ProviderName = 'Service Control Manager' } -MaxEvents $EventLogMax -ErrorAction SilentlyContinue |
                    Where-Object { $_.Message -match 'Spooler|Print' } |
                    Select-Object TimeCreated, Id, LevelDisplayName, @{ n = 'Message'; e = { ($_.Message -replace '\s+', ' ') } } |
                    Select-Object -First 60)
            } catch { $d['EventLog'] = @() }
            try {
                $d['PrintEventLog'] = @(Get-WinEvent -LogName 'Microsoft-Windows-PrintService/Admin' -MaxEvents 60 -ErrorAction SilentlyContinue |
                    Where-Object { $_.TimeCreated -ge $since } |
                    Select-Object TimeCreated, Id, LevelDisplayName, @{ n = 'Message'; e = { ($_.Message -replace '\s+', ' ') } })
            } catch { $d['PrintEventLog'] = @() }
        }
        return $d
    }

    # ------------------------------------------------------------------------
    #  MAIN
    # ------------------------------------------------------------------------
    $mode = [string]$Options['Mode']
    $final = 'Failed'
    $R['Mode'] = $mode

    try {
        # -- resolve + guard the spool directory -----------------------------
        $resolved = Resolve-SpoolDir
        $R['SpoolDirectory']       = $resolved.Path
        $R['SpoolDirectorySource'] = $resolved.Source
        $guardError = Test-SpoolDirSafe -dir $resolved.Path
        if ($guardError) {
            Add-E $guardError
            $R['Status'] = 'Failed'
            $R['SpoolDirectoryValid'] = $false
            throw $guardError
        }
        $R['SpoolDirectoryValid'] = $true
        if (-not (Test-Path -LiteralPath $R['SpoolDirectory'])) {
            Add-E ("Spool directory does not exist: {0}" -f $R['SpoolDirectory'])
            $R['Status'] = 'Failed'; throw 'Spool directory missing'
        }

        # -- diagnostics mode: read-only, never touches the service ----------
        if ($mode -eq 'Diagnostic') {
            $R['Diagnostics'] = Get-DiagBundle -Dir $R['SpoolDirectory'] -EventLogHours ([int]$Options['EventLogHours']) -EventLogMax ([int]$Options['EventLogMaxEntries'])
            $R['DiagnosticsCollected'] = $true
            $R['Jobs'] = @(Get-PrintJobsWmi)
            if (@($R['Jobs']).Count -eq 0) { $R['Jobs'] = @(Get-PrintJobsFromSpoolDir -Dir $R['SpoolDirectory']) }
            $R['JobsFound'] = @($R['Jobs']).Count
            $svc = Get-SvcState $spoolSvc
            $R['ServiceWasRunning'] = ($svc.State -eq 'Running')
            $R['ServiceIsRunning']  = ($svc.State -eq 'Running')
            return Finish-Result 'Diagnostic'
        }

        # -- inventory BEFORE stopping (WMI needs a live spooler) ------------
        $svcBefore = Get-SvcState $spoolSvc
        if (-not $svcBefore.Exists) { Add-E ("Service '{0}' is not installed on this host." -f $spoolSvc); $R['Status'] = 'Failed'; throw 'No spooler service' }
        $R['ServiceWasRunning'] = ($svcBefore.State -eq 'Running')

        # ---- inventory: WMI (rich metadata) correlated with .SHD (file truth)
        # WMI gives printer/document/owner/size but NOT the file names; the
        # .SHD scan gives exact file names but only best-effort metadata. We
        # merge them so that selective deletion can be both precise and safe.
        $jobs = @()
        if ($R['ServiceWasRunning']) { $jobs = @(Get-PrintJobsWmi) }
        $dirJobs = @(Get-PrintJobsFromSpoolDir -Dir $R['SpoolDirectory'])

        if ($jobs.Count -eq 0) {
            $jobs = $dirJobs
            if ($dirJobs.Count -gt 0) { Add-W 'The Spooler did not answer a WMI job query (typical when it is hung); the inventory was built from .SHD headers instead. Printer/document names are best-effort in this mode.' }
        }
        else {
            # Correlate by ascending order. The spooler assigns .SHD/.SPL file
            # numbers monotonically in submission order, and Win32_PrintJob
            # enumerates in the same order, so positional pairing is reliable
            # when the counts agree. When they disagree we pair what we can and
            # report the rest as orphans rather than guessing.
            $wmiSorted = @($jobs | Sort-Object { [long]$_.JobId })
            $shdSorted = @($dirJobs | Sort-Object { [long]$_.JobId })
            $pairCount = [Math]::Min($wmiSorted.Count, $shdSorted.Count)
            for ($i = 0; $i -lt $pairCount; $i++) {
                $wmiSorted[$i] = [pscustomobject]@{
                    JobId        = $wmiSorted[$i].JobId
                    PrinterName  = $(if ($wmiSorted[$i].PrinterName) { $wmiSorted[$i].PrinterName } else { $shdSorted[$i].PrinterName })
                    DocumentName = $(if ($wmiSorted[$i].DocumentName) { $wmiSorted[$i].DocumentName } else { $shdSorted[$i].DocumentName })
                    UserName     = $(if ($wmiSorted[$i].UserName) { $wmiSorted[$i].UserName } else { $shdSorted[$i].UserName })
                    SizeBytes    = $wmiSorted[$i].SizeBytes
                    TotalPages   = $wmiSorted[$i].TotalPages
                    PagesPrinted = $wmiSorted[$i].PagesPrinted
                    Status       = $wmiSorted[$i].Status
                    Priority     = $wmiSorted[$i].Priority
                    SubmittedUtc = $wmiSorted[$i].SubmittedUtc
                    Source       = 'WMI+SHD'
                    ShdFile      = $shdSorted[$i].ShdFile
                    SplFile      = $shdSorted[$i].SplFile
                }
            }
            $jobs = @($wmiSorted)
            if ($shdSorted.Count -gt $pairCount) {
                $orphans = @($shdSorted[$pairCount..($shdSorted.Count - 1)])
                Add-W ("{0} spool file pair(s) exist without a matching WMI job (orphaned/leftover). They are included in the inventory as JobId from the file name." -f $orphans.Count)
                foreach ($o in $orphans) {
                    $jobs += $o
                    if ($o.ShdFile) { $R['OrphanedFiles'].Add($o.ShdFile) | Out-Null }
                    if ($o.SplFile) { $R['OrphanedFiles'].Add($o.SplFile) | Out-Null }
                }
            }
            elseif ($wmiSorted.Count -gt $pairCount) {
                Add-W ("{0} WMI job(s) have no .SHD file on disk; they cannot be deleted as files and will be dropped when the service restarts." -f ($wmiSorted.Count - $pairCount))
            }
        }
        $R['JobsFound'] = @($jobs).Count
        # PagesPrinted and Priority are computed above (both the WMI path and the
        # .SHD fallback set them) and are documented in Report-Schema.md, so they
        # MUST survive this projection. An earlier revision listed every field
        # except these two, which silently dropped them from the JSON contract -
        # the schema promised columns no consumer would ever receive.
        $R['Jobs'] = @($jobs | Select-Object JobId, PrinterName, DocumentName, UserName, SizeBytes, TotalPages, PagesPrinted, Status, Priority, SubmittedUtc, Source, ShdFile, SplFile)

        # -- file inventory (everything physically present) ------------------
        $allFiles = @(Get-ChildItem -LiteralPath $R['SpoolDirectory'] -File -Force -ErrorAction SilentlyContinue)
        $spoolFiles = @($allFiles | Where-Object { $exts -contains $_.Extension.ToUpperInvariant() })
        $R['FilesFound'] = $spoolFiles.Count
        $bytesFound = [long]0
        foreach ($bf in $spoolFiles) { $bytesFound += [long]$bf.Length }
        $R['BytesFound'] = $bytesFound
        if ($R['FilesFound'] -gt @@GUARD_MAX_FILES@@) {
            Add-E ("Refusing to continue: {0} spool files found (guard limit @@GUARD_MAX_FILES@@). Investigate with -Diagnostic." -f $R['FilesFound'])
            $R['Status'] = 'Failed'; throw 'Too many spool files'
        }
        if ($R['BytesFound'] -gt @@GUARD_MAX_BYTES@@) {
            Add-E ("Refusing to continue: {0} bytes of spool data (guard limit @@GUARD_MAX_BYTES@@)." -f $R['BytesFound'])
            $R['Status'] = 'Failed'; throw 'Spool data too large'
        }

        # -- selection -------------------------------------------------------
        $selection = Select-JobsToActOn -Jobs $jobs
        if ($selection.Note) { Add-W $selection.Note }
        $selectedJobs = @($selection.Selected)
        $R['JobsSelected'] = $selectedJobs.Count

        $targetFiles = New-Object System.Collections.Generic.List[string]
        if ($selection.Selective) {
            foreach ($j in $selectedJobs) { foreach ($f in (Get-FilesForJob -Job $j -Dir $R['SpoolDirectory'])) { if (-not $targetFiles.Contains($f)) { $targetFiles.Add($f) | Out-Null } } }
        } else {
            foreach ($f in $spoolFiles) { $targetFiles.Add($f.FullName) | Out-Null }
            if ($Options['PurgeLegacy']) {
                foreach ($f in @($allFiles | Where-Object { $legacyExt -contains $_.Extension.ToUpperInvariant() -or $_.Name -like '~*' })) {
                    if (-not $targetFiles.Contains($f.FullName)) { $targetFiles.Add($f.FullName) | Out-Null }
                }
            }
        }

        $R['FilesSelected'] = $targetFiles.Count

        if ($R['FilesSelected'] -eq 0) {
            if ($R['FilesFound'] -gt 0) { Add-W ("{0} spool file(s) exist but none matched the selection filters; nothing was deleted." -f $R['FilesFound']) }
            $R['ServiceIsRunning'] = $R['ServiceWasRunning']
            return Finish-Result 'NothingToDo'
        }

        # -- dry run stops here ---------------------------------------------
        if ($R['DryRun']) {
            foreach ($f in $targetFiles) {
                $sz = [long]0
                try { $sz = (Get-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue).Length } catch { /* size probing is informational only */ }
                $R['RemovedFiles'].Add([pscustomobject]@{ File = [System.IO.Path]::GetFileName($f); Bytes = $sz; DryRun = $true }) | Out-Null
                $R['BytesDeleted'] += $sz
            }
            $R['JobsRemoved'] = $selectedJobs.Count
            return Finish-Result 'DryRun'
        }

        # -- quiescence wait (avoid deleting jobs that are actively printing)
        if ($Options['QuiescenceSeconds'] -gt 0 -and $R['ServiceWasRunning']) {
            $qDeadline = (Get-Date).AddSeconds([int]$Options['QuiescenceSeconds'])
            $lastCount = -1; $stableSince = Get-Date
            while ((Get-Date) -lt $qDeadline) {
                $c = @(Get-ChildItem -LiteralPath $R['SpoolDirectory'] -File -Force -ErrorAction SilentlyContinue | Where-Object { $exts -contains $_.Extension.ToUpperInvariant() }).Count
                if ($c -ne $lastCount) { $lastCount = $c; $stableSince = Get-Date }
                elseif (((Get-Date) - $stableSince).TotalSeconds -ge [Math]::Min(5, $Options['QuiescenceSeconds'])) { break }
                Start-Sleep -Milliseconds 500
            }
        }

        # -- backup BEFORE anything destructive ------------------------------
        if ($Options['Backup']) {
            $bk = New-SpoolBackup -Files $targetFiles.ToArray() -BackupRoot $Options['BackupPath'] -Computer $env:COMPUTERNAME -SelectedJobs $selectedJobs
            if ($bk.Error) {
                Add-E ("Backup failed: {0}" -f $bk.Error)
                if ($Options['BackupRequired']) { $R['Status'] = 'Failed'; throw 'Backup required but failed' }
            } else {
                $R['BackupPath'] = $bk.Path; $R['BackupHash'] = $bk.Hash
                if ($bk.FilesCopied -lt $targetFiles.Count) { Add-W ("Backup captured {0} of {1} selected file(s)." -f $bk.FilesCopied, $targetFiles.Count) }
            }
        }

        # -- stop dependent services, then the Spooler -----------------------
        $stopSw = [System.Diagnostics.Stopwatch]::StartNew()
        $deps = @()
        $depsToRestart = New-Object System.Collections.Generic.List[string]

        if ($Options['StopDependents']) {
            $deps = @(Resolve-DependentServiceNames)
            foreach ($d in $deps) {
                $st = Get-SvcState $d
                $res = Stop-OneService -name $d -timeoutSec ([int]$Options['ServiceTimeoutSeconds']) -AllowKill:$false
                if ($st.State -eq 'Running') { $depsToRestart.Add($d) | Out-Null }
                if (-not $res.Ok) { Add-W ("Dependent service '{0}' did not stop cleanly (state: {1})." -f $d, $res.State) }
            }
        }

        $stopRes = Stop-OneService -name $spoolSvc -timeoutSec ([int]$Options['ServiceTimeoutSeconds']) -AllowKill:([bool]$Options['ForceStopFallback'])
        $stopSw.Stop()
        $R['ServiceStopMs'] = $stopSw.ElapsedMilliseconds

        if (-not $stopRes.Ok) {
            Add-E ("Print Spooler could not be stopped (state: {0}). No files were deleted." -f $stopRes.State)
            # Make sure we do not leave dependents down.
            foreach ($d in $depsToRestart) { $null = Start-OneService -name $d -timeoutSec ([int]$Options['ServiceTimeoutSeconds']) }
            $null = Start-OneService -name $spoolSvc -timeoutSec ([int]$Options['ServiceTimeoutSeconds'])
            $svcAfter = Get-SvcState $spoolSvc
            $R['ServiceIsRunning'] = ($svcAfter.State -eq 'Running')
            $R['DependentServices'] = @($depsToRestart.ToArray())
            return Finish-Result 'ServiceFailure'
        }

        if ($Options['SettleSeconds'] -gt 0) { Start-Sleep -Seconds ([int]$Options['SettleSeconds']) }

        # -- delete ----------------------------------------------------------
        $deletedBytes = [long]0
        foreach ($f in $targetFiles.ToArray()) {
            $fname = [System.IO.Path]::GetFileName($f)
            $size = [long]0
            try { $size = (Get-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue).Length } catch { /* size probing is informational only */ }
            $res = Remove-SpoolFileRobust -Path $f -RetryCount ([int]$Options['RetryCount']) -RetryDelayMs ([int]$Options['RetryDelayMs'])
            if ($res.Ok) {
                $R['FilesDeleted']++
                $deletedBytes += $size
                $R['RemovedFiles'].Add([pscustomobject]@{ File = $fname; Bytes = $size; Attempts = $res.Attempts }) | Out-Null
            } else {
                $R['FilesFailed']++
                if ($res.Locked) { $R['FilesLocked']++ }
                $R['FailedFiles'].Add([pscustomobject]@{ File = $fname; Bytes = $size; Locked = $res.Locked; Attempts = $res.Attempts; Error = $res.Error }) | Out-Null
                Add-W ("Could not delete '{0}' (locked={1}, attempts={2}): {3}" -f $fname, $res.Locked, $res.Attempts, $res.Error)
            }
        }
        $R['BytesDeleted'] = $deletedBytes

        # -- optional deep repair (finding A6) -------------------------------
        if ($Options['Repair']) {
            try {
                if ($Options['PurgeFormsCache']) {
                    $formsDir = Join-Path $env:SystemRoot '@@FORMS_SUBPATH@@'
                    if (Test-Path -LiteralPath $formsDir) {
                        $forms = @(Get-ChildItem -LiteralPath $formsDir -File -Force -ErrorAction SilentlyContinue)
                        $killed = 0
                        foreach ($fm in $forms) {
                            try { [System.IO.File]::Delete($fm.FullName); $killed++ }
                            catch { /* one undeletable form must not stop the rest */ }
                        }
                        Add-W ("Repair: cleared {0} of {1} Forms cache file(s) in {2}." -f $killed, $forms.Count, $formsDir)
                    }
                }
                if ($Options['ResetSpoolDirectoryValue']) {
                    try {
                        Remove-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers' -Name 'DefaultSpoolDirectory' -ErrorAction Stop
                        Add-W 'Repair: removed the DefaultSpoolDirectory override; the OS default will be used.'
                    }
                    catch { Add-W ("Repair: could not reset DefaultSpoolDirectory: {0}" -f $_.Exception.Message) }
                }
                if ($Options['PurgeDriverStore']) {
                    try {
                        $unused = @(Get-PrintDriver -ErrorAction Stop | Where-Object { -not (Get-Printer -DriverName $_.Name -ErrorAction SilentlyContinue) })
                        $removed = 0
                        foreach ($drv in $unused) {
                            try { Remove-PrintDriver -Name $drv.Name -ErrorAction Stop; $removed++ }
                            catch { /* a driver still in use is skipped, not fatal */ }
                        }
                        Add-W ("Repair: removed {0} unreferenced printer driver package(s) from the driver store." -f $removed)
                    }
                    catch { Add-W ("Repair: driver store purge unavailable: {0}" -f $_.Exception.Message) }
                }
                $R['RepairPerformed'] = $true
            }
            catch { Add-E ("Repair step failed: {0}" -f $_.Exception.Message) }
        }

        # -- restart Spooler first, then dependents --------------------------
        $startSw = [System.Diagnostics.Stopwatch]::StartNew()
        $startOk = Start-OneService -name $spoolSvc -timeoutSec ([int]$Options['ServiceTimeoutSeconds'])
        if ($Options['RestartDependents']) {
            foreach ($d in $depsToRestart) {
                $ok = Start-OneService -name $d -timeoutSec ([int]$Options['ServiceTimeoutSeconds'])
                if (-not $ok) { Add-W ("Dependent service '{0}' failed to restart." -f $d) }
            }
        }
        $startSw.Stop()
        $R['ServiceStartMs'] = $startSw.ElapsedMilliseconds
        $R['DependentServices'] = @($depsToRestart.ToArray() | ForEach-Object { $_ })

        $svcAfter = Get-SvcState $spoolSvc
        $R['ServiceIsRunning'] = ($svcAfter.State -eq 'Running')

        if (-not $startOk -or -not $R['ServiceIsRunning']) {
            Add-E ("Print Spooler did not reach RUNNING within {0}s (state: {1}). Printing is unavailable until the service is started." -f $Options['ServiceTimeoutSeconds'], $svcAfter.State)
            return Finish-Result 'ServiceFailure'
        }

        # -- post-restart verification: did the spooler resurrect files? -----
        if ($Options['SettleSeconds'] -gt 0) { Start-Sleep -Seconds ([Math]::Min(5, [int]$Options['SettleSeconds'])) }
        $after = @(Get-ChildItem -LiteralPath $R['SpoolDirectory'] -File -Force -ErrorAction SilentlyContinue | Where-Object { $exts -contains $_.Extension.ToUpperInvariant() })
        $R['FilesAfter'] = $after.Count
        if ($Options['VerifyAfterRestart'] -and -not $selection.Selective -and $after.Count -gt 0) {
            $R['ResurrectedFiles'] = $after.Count
            Add-W ("{0} spool file(s) reappeared after the service restart. A driver or port monitor is re-submitting them; run -Diagnostic and check the driver/port." -f $after.Count)
        }

        # -- final status ----------------------------------------------------
        $final = 'Success'
        if ($R['FilesDeleted'] -eq 0 -and $R['FilesSelected'] -gt 0) { $final = 'Failed' }
        elseif ($R['FilesFailed'] -gt 0 -or $R['ResurrectedFiles'] -gt 0 -or $R['Warnings'].Count -gt 0) { $final = 'SuccessWithWarnings' }
        $R['JobsRemoved'] = $selectedJobs.Count
    }
    catch {
        $msg = $_.Exception.Message
        Add-E ("Unhandled payload error: {0}" -f $msg)
        # LAST RESORT: never leave the machine without a spooler. A cleanup that
        # fails halfway must still hand back a host that can print.
        try { $null = Start-OneService -name $spoolSvc -timeoutSec 30 }
        catch { /* the service may already be running; nothing further is safe to try */ }
        try { $R['ServiceIsRunning'] = ((Get-SvcState $spoolSvc).State -eq 'Running') }
        catch { /* state reporting is best-effort inside the failure path */ }
        return Finish-Result 'Failed'
    }

    return Finish-Result $final
'@

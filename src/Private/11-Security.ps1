# ============================================================================
#  Elevation, code-signing verification, integrity, diagnostics bundle
# ----------------------------------------------------------------------------
#  FINDING A22: v3 required the operator to notice a failed admin check and
#  manually re-launch. v4 can self-elevate (with consent), detects a
#  non-interactive host and fails fast instead of hanging, and can verify its own
#  Authenticode signature before it touches anything.
# ============================================================================

function Invoke-SelfElevation {
    <#
    .SYNOPSIS
        Re-launches the current invocation in an elevated process.
    .DESCRIPTION
        Rebuilds the command line from the script path plus every bound argument,
        then starts it with ShellExecute "runas". Appends -NoElevate to the child
        so an elevation failure can never loop.
    .OUTPUTS
        PSCustomObject { Attempted, Success, ExitCode, Message, CommandLine }
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyString()][string]$ScriptPath,
        [AllowNull()][string[]]$ChildArguments,
        [switch]$PassThruOnly
    )

    $out = [ordered]@{
        Attempted   = $false
        Success     = $false
        ExitCode    = $null
        Message     = $null
        CommandLine = $null
    }

    if (-not (Test-RunningOnWindows)) {
        $out.Message = 'Self-elevation is only supported on Windows.'
        return [pscustomobject]$out
    }
    if (Test-RunningAsAdministrator) {
        $out.Success = $true
        $out.Message = 'Already running with administrator privileges.'
        return [pscustomobject]$out
    }
    if (Test-NonInteractiveHost) {
        $out.Message = 'Cannot elevate from a non-interactive host (Task Scheduler / CI). Run the task as SYSTEM or an elevated account instead.'
        return [pscustomobject]$out
    }
    if ([string]::IsNullOrWhiteSpace($ScriptPath)) {
        $out.Message = 'Cannot elevate: the script path is unknown (the script was probably pasted into the console instead of being run from a file).'
        return [pscustomobject]$out
    }

    $out.Attempted = $true

    $quoted = New-Object System.Collections.Generic.List[string]
    foreach ($a in @($ChildArguments)) {
        if ([string]::IsNullOrWhiteSpace($a)) { continue }
        if ($a -match '[\s"]') { $quoted.Add('"{0}"' -f ($a -replace '"', '\"')) }
        else { $quoted.Add($a) }
    }
    if (@($quoted) -notcontains '-NoElevate') { $quoted.Add('-NoElevate') }

    # Re-launch under the same engine we are running in, so a PowerShell 7 user is
    # not silently downgraded to Windows PowerShell 5.1 by the UAC relaunch.
    $exe = $null
    try { $exe = (Get-Process -Id $PID -ErrorAction Stop).Path }
    catch { $exe = $null }
    if ([string]::IsNullOrWhiteSpace($exe)) { $exe = 'powershell.exe' }

    $argLine = '-NoProfile -ExecutionPolicy Bypass -File "{0}" {1}' -f $ScriptPath, ($quoted -join ' ')
    $out.CommandLine = "$exe $argLine"

    if ($PassThruOnly) { return [pscustomobject]$out }

    $proc = $null
    try {
        Write-LogWarn 'Administrator privileges are required. Requesting elevation (UAC prompt)...' -Component 'Security'
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $exe
        $psi.Arguments = $argLine
        $psi.UseShellExecute = $true
        $psi.Verb = 'runas'
        $psi.WorkingDirectory = [Environment]::CurrentDirectory

        $proc = [System.Diagnostics.Process]::Start($psi)
        if ($proc) {
            $proc.WaitForExit()
            $out.ExitCode = $proc.ExitCode
            $out.Success = $true
            $out.Message = "The elevated process finished with exit code $($proc.ExitCode)."
        }
        else {
            $out.Message = 'Elevation was started but no process handle was returned.'
        }
    }
    catch {
        if ($_.Exception.Message -match 'canceled|cancelled|The operation was canceled by the user') {
            $out.Message = 'Elevation was declined by the user (UAC cancelled).'
            $out.ExitCode = $Script:ExitCode.Cancelled
        }
        else {
            $out.Message = "Elevation failed: $($_.Exception.Message)"
        }
    }
    finally {
        if ($proc) { $proc.Dispose() }
    }
    return [pscustomobject]$out
}

function Test-ScriptSignature {
    <#
    .SYNOPSIS
        Authenticode verification of a file.
    .DESCRIPTION
        Returns a status object rather than throwing, because "not signed" is a
        perfectly normal state for an open-source tool and must not be confused
        with "the signature is broken".
        Status values:
          Valid              signed, chain trusted, and any expected thumbprint matches
          NotSigned          no signature present
          Invalid            signature present but untrusted, or the hash does not match
          ThumbprintMismatch trusted signature from an unexpected publisher
          NotSupported       this platform cannot verify Authenticode
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)][AllowEmptyString()][string]$Path,
        [AllowNull()][AllowEmptyString()][string]$ExpectedThumbprint,
        [AllowNull()][AllowEmptyString()][string]$ExpectedSubject
    )

    $res = [ordered]@{
        Path          = $Path
        Status        = 'NotSupported'
        SignerSubject = $null
        Thumbprint    = $null
        SignatureType = $null
        Message       = $null
        VerifiedUtc   = [DateTime]::UtcNow.ToString('o')
    }

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $res.Status = 'Invalid'
        $res.Message = "File not found: $Path"
        return [pscustomobject]$res
    }
    if (-not (Get-Command -Name Get-AuthenticodeSignature -ErrorAction SilentlyContinue)) {
        $res.Message = 'Get-AuthenticodeSignature is unavailable on this platform.'
        return [pscustomobject]$res
    }

    try {
        $sig = Get-AuthenticodeSignature -LiteralPath $Path -ErrorAction Stop
        $res.SignatureType = [string]$sig.SignatureType
        $res.Message = [string]$sig.StatusMessage
        $res.Status = switch ([string]$sig.Status) {
            'Valid' { 'Valid' }
            'NotSigned' { 'NotSigned' }
            default { 'Invalid' }
        }
        if ($sig.SignerCertificate) {
            $res.SignerSubject = $sig.SignerCertificate.Subject
            $res.Thumbprint = $sig.SignerCertificate.Thumbprint
        }

        if ($res.Status -eq 'Valid' -and $ExpectedThumbprint) {
            if (-not (Test-HashMatch -Expected $ExpectedThumbprint -Actual $res.Thumbprint)) {
                $res.Status = 'ThumbprintMismatch'
                $res.Message = "The signature is valid but was published with an unexpected certificate. Expected thumbprint '$ExpectedThumbprint', found '$($res.Thumbprint)'."
            }
        }
        if ($res.Status -eq 'Valid' -and $ExpectedSubject) {
            if ($res.SignerSubject -notlike "*$ExpectedSubject*") {
                $res.Status = 'ThumbprintMismatch'
                $res.Message = "The signature is valid but the subject '$($res.SignerSubject)' does not match the expected '$ExpectedSubject'."
            }
        }
    }
    catch {
        $res.Status = 'Invalid'
        $res.Message = "Signature verification failed: $($_.Exception.Message)"
    }
    return [pscustomobject]$res
}

function Get-FileIntegrityRecord {
    <#
    .SYNOPSIS
        Size, SHA-256 and Authenticode status of a file, for release and audit records.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory, Position = 0)][AllowEmptyString()][string]$Path)

    $rec = [ordered]@{
        Path         = $Path
        Exists       = $false
        SizeBytes    = [long]0
        Sha256       = $null
        LastWriteUtc = $null
        Signature    = $null
    }
    if (-not [string]::IsNullOrWhiteSpace($Path) -and (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $item = Get-Item -LiteralPath $Path -Force
        $rec.Exists = $true
        $rec.SizeBytes = $item.Length
        $rec.LastWriteUtc = $item.LastWriteTimeUtc.ToString('o')
        $rec.Sha256 = Get-FileHashSafe -Path $Path -Algorithm 'SHA256'
        $sig = Test-ScriptSignature -Path $Path
        $rec.Signature = [pscustomobject]@{ Status = $sig.Status; Subject = $sig.SignerSubject; Thumbprint = $sig.Thumbprint }
    }
    return [pscustomobject]$rec
}

function Write-DiagnosticsBundle {
    <#
    .SYNOPSIS
        Produces a single ZIP a user can attach to a GitHub issue.
    .DESCRIPTION
        Contains the environment report, the diagnostic payload output, the recent
        log ring buffer, the effective configuration with secrets redacted, and the
        integrity records of the installed files. This is the difference between
        "it does not work" and a report a maintainer can act on.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyString()][string]$DestinationDirectory,
        [AllowNull()]$Summary,
        [AllowNull()][System.Collections.IDictionary]$Configuration
    )

    $stamp = Get-SafeFileNameStamp
    if (-not [string]::IsNullOrWhiteSpace($DestinationDirectory)) {
        $root = Resolve-DirectoryPath -Path $DestinationDirectory -Create
    }
    else {
        $root = Resolve-DirectoryPath -Path (Join-Path ([IO.Path]::GetTempPath()) $Script:ProductFamily) -Create
    }

    $name = "{0}-diagnostics_{1}" -f $Script:ScriptName, $stamp
    $stage = Join-Path $root $name
    if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue }
    New-Item -ItemType Directory -Path $stage -Force -ErrorAction Stop | Out-Null

    try {
        $envReport = [ordered]@{
            Tool              = $Script:ScriptName
            Version           = $Script:ScriptVersion
            Build             = $Script:ScriptBuild
            GeneratedUtc      = [DateTime]::UtcNow.ToString('o')
            PowerShellVersion = $PSVersionTable.PSVersion.ToString()
            PSEdition         = $PSVersionTable.PSEdition
            CLRVersion        = $PSVersionTable.CLRVersion
            OS                = (ConvertTo-HashtableDeep -InputObject (Get-WindowsSku))
            IsWindows         = (Test-RunningOnWindows)
            IsAdministrator   = (Test-RunningAsAdministrator)
            Context           = (ConvertTo-HashtableDeep -InputObject (Get-CurrentUserContext))
            ConsoleCapability = (ConvertTo-HashtableDeep -InputObject (Get-ConsoleCapability))
            Culture           = (Get-Culture).Name
            UiCulture         = (Get-UICulture).Name
            InstallDirectory  = $Script:RuntimeState.InstallDirectory
            LogPath           = $Script:RuntimeState.LogPath
            ModulesLoaded     = @(Get-Module | Select-Object -ExpandProperty Name)
        }
        ($envReport | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath (Join-Path $stage 'environment.json') -Encoding UTF8

        if ($Summary) {
            (ConvertTo-JsonReport -Summary $Summary) | Set-Content -LiteralPath (Join-Path $stage 'run-report.json') -Encoding UTF8
            (ConvertTo-TextReport -Summary $Summary) | Set-Content -LiteralPath (Join-Path $stage 'run-report.txt') -Encoding UTF8
        }

        $buffer = @(Get-LogBufferSnapshot)
        if ($buffer.Count -gt 0) {
            $text = ($buffer | ForEach-Object { "{0}|{1}|{2}|{3}|{4}" -f $_.TimeUtc.ToString('o'), $_.Level, $_.Component, $_.Target, $_.Message }) -join [Environment]::NewLine
            Set-Content -LiteralPath (Join-Path $stage 'log-buffer.txt') -Value $text -Encoding UTF8
        }

        if ($Configuration) {
            $redacted = ConvertTo-HashtableDeep -InputObject $Configuration
            if (($redacted -is [System.Collections.IDictionary]) -and $redacted.Contains('Update')) {
                $upd = $redacted['Update']
                if (($upd -is [System.Collections.IDictionary]) -and $upd.Contains('Token') -and $upd['Token']) {
                    # A GitHub token must never leave the machine in a bug report.
                    $upd['Token'] = '<redacted>'
                }
            }
            ($redacted | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath (Join-Path $stage 'effective-config.json') -Encoding UTF8
        }

        $integrity = New-Object System.Collections.Generic.List[object]
        foreach ($f in @('Clear-PrintSpoolerQueue.ps1', 'Clear-PrintSpoolerQueue.bat')) {
            $p = Join-Path $Script:RuntimeState.InstallDirectory $f
            if (Test-Path -LiteralPath $p) { $integrity.Add((ConvertTo-HashtableDeep -InputObject (Get-FileIntegrityRecord -Path $p))) }
        }
        ($integrity.ToArray() | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath (Join-Path $stage 'file-integrity.json') -Encoding UTF8

        if ($Script:RuntimeState.LogPath -and (Test-Path -LiteralPath $Script:RuntimeState.LogPath)) {
            try { Copy-Item -LiteralPath $Script:RuntimeState.LogPath -Destination (Join-Path $stage 'current.log') -Force -ErrorAction Stop }
            catch { /* the log copy is a convenience; the buffer above already carries the recent entries */ }
        }

        $zip = Join-Path $root ("$name.zip")
        if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force -ErrorAction SilentlyContinue }
        if (Get-Command -Name Compress-Archive -ErrorAction SilentlyContinue) {
            Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip -CompressionLevel Optimal -ErrorAction Stop
        }
        else {
            Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
            [System.IO.Compression.ZipFile]::CreateFromDirectory($stage, $zip)
        }
        Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue
        return $zip
    }
    catch {
        Write-LogError "Could not build the diagnostics bundle: $($_.Exception.Message)" -Component 'Diagnostics'
        return $stage
    }
}

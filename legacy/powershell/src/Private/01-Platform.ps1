# ============================================================================
#  Platform, host-capability and environment detection
# ----------------------------------------------------------------------------
#  Everything in this file must be safe to call on PowerShell 5.1 AND 7.x, on
#  Windows AND (for -CheckUpdate / -Help only) on Linux/macOS. The automatic
#  variables $IsWindows/$IsLinux/$IsMacOS do NOT exist on Windows PowerShell
#  5.1, so they are probed with Get-Variable instead of being referenced
#  directly - a direct reference under Set-StrictMode throws.
# ============================================================================

function Test-RunningOnWindows {
    <#
    .SYNOPSIS
        Returns $true when the host OS is Windows (any SKU, any PowerShell edition).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    $v = Get-Variable -Name 'IsWindows' -Scope Global -ErrorAction SilentlyContinue
    if ($null -ne $v) { return [bool]$v.Value }
    # Windows PowerShell 5.1 and earlier define no $IsWindows variable at all.
    return ($PSVersionTable.PSVersion.Major -le 5) -or ($env:OS -eq 'Windows_NT')
}

function Test-RunningAsAdministrator {
    <#
    .SYNOPSIS
        True when the current token carries the Administrators role (elevated).
    .NOTES
        Replaces the v3 "net session" probe used by the batch launcher: that test
        also fails on domain controllers with restricted session rights and on
        machines where the Server service is disabled, even when the token IS
        elevated. The principal check is authoritative.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    if (-not (Test-RunningOnWindows)) { return $false }
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        # An unreadable token must not be reported as "elevated".
        return $false
    }
}

function Get-CurrentUserContext {
    <#
    .SYNOPSIS
        Returns a PSCustomObject describing the security and execution context.
    #>
    [CmdletBinding()]
    param()

    $ctx = [ordered]@{
        UserName      = $null
        IsAdmin       = $false
        IsSystem      = $false
        IsInteractive = $false
        SessionType   = 'Unknown'
    }
    try {
        if (Test-RunningOnWindows) {
            $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
            $ctx.UserName = $identity.Name
            $ctx.IsSystem = ($identity.IsSystem -or $ctx.UserName -eq 'NT AUTHORITY\SYSTEM')
            $ctx.IsAdmin = Test-RunningAsAdministrator
        }
        else {
            $ctx.UserName = $env:USER
        }
        $ctx.IsInteractive = [Environment]::UserInteractive -and -not [Console]::IsInputRedirected
        if ($ctx.IsSystem) { $ctx.SessionType = 'Service/SYSTEM' }
        elseif ($ctx.IsInteractive) { $ctx.SessionType = 'Interactive' }
        else { $ctx.SessionType = 'Non-interactive/Batch' }
    }
    catch {
        # Context probing is informational; it must never break a cleanup run.
    }
    return [pscustomobject]$ctx
}

function Test-NonInteractiveHost {
    <#
    .SYNOPSIS
        True when prompting the user is impossible (Task Scheduler, CI, redirected stdin).
    .NOTES
        FINDING A22: v3 hung forever on "Press any key" / `set /p` when launched
        from Task Scheduler. v4 detects the situation and refuses to prompt,
        failing fast with a clear message instead of blocking a maintenance window.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    if ($Script:RuntimeState.NonInteractive) { return $true }
    try { if ([Environment]::UserInteractive -eq $false) { return $true } } catch { return $true }
    try { if ([Console]::IsInputRedirected) { return $true } } catch { return $true }
    return $false
}

function Test-PowerShellSevenFeatures {
    <#
    .SYNOPSIS
        True when ForEach-Object -Parallel is available (PowerShell 7.0+).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    return ($PSVersionTable.PSVersion -ge [Version]'7.0')
}

function Get-ConsoleCapability {
    <#
    .SYNOPSIS
        Detects colour, Unicode-glyph and width support of the current host.
    .DESCRIPTION
        Drives the automatic ASCII/Unicode and colour decisions so redirected logs
        and legacy OEM code pages (437/850) never contain mojibake. v3 emitted raw
        Unicode glyphs into cmd.exe's OEM code page and produced garbage.
    #>
    [CmdletBinding()]
    param()

    $cap = [ordered]@{
        SupportsColor     = $true
        SupportsUnicode   = $true
        OutputRedirected  = $false
        CodePage          = 0
        BufferWidth       = 120
        IsWindowsTerminal = $false
    }

    try { $cap.OutputRedirected = [bool][Console]::IsOutputRedirected }
    catch { $cap.OutputRedirected = $true }

    try {
        if (Test-RunningOnWindows) {
            $cap.CodePage = [Console]::OutputEncoding.CodePage
            # 65001 = UTF-8, 1200 = UTF-16. Anything else is a legacy OEM/ANSI
            # page that cannot render U+2714 and friends.
            $cap.SupportsUnicode = ($cap.CodePage -eq 65001) -or ($cap.CodePage -eq 1200)
        }
        else {
            $cap.SupportsUnicode = $true
        }
    }
    catch {
        $cap.SupportsUnicode = $false
    }

    try { $cap.BufferWidth = [Math]::Max(60, [int]$Host.UI.RawUI.BufferSize.Width) }
    catch {
        try { $cap.BufferWidth = [Math]::Max(60, [int][Console]::WindowWidth) }
        catch { $cap.BufferWidth = 120 }
    }

    if ($cap.OutputRedirected) { $cap.SupportsColor = $false; $cap.SupportsUnicode = $false }
    if ($env:WT_SESSION -or $env:TERMINAL_EMULATOR) { $cap.IsWindowsTerminal = $true }
    if ($env:NO_COLOR) { $cap.SupportsColor = $false }   # https://no-color.org

    return [pscustomobject]$cap
}

function Get-SpoolGuardForbiddenRoots {
    <#
    .SYNOPSIS
        Forbidden parent directories for the delete step, resolved at runtime.
    .DESCRIPTION
        The static literal list lives in the constants file; machine-specific
        roots are added here so an unset environment variable can never inject
        $null into the comparison set.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param()

    $roots = New-Object System.Collections.Generic.List[string]
    foreach ($static in $Script:SpoolGuardForbiddenRoots) { [void]$roots.Add($static) }

    foreach ($varName in @('SystemRoot', 'windir', 'ProgramFiles', 'ProgramFiles(x86)', 'ProgramData', 'USERPROFILE', 'SystemDrive')) {
        $val = [Environment]::GetEnvironmentVariable($varName)
        if (-not [string]::IsNullOrWhiteSpace($val)) {
            [void]$roots.Add($val)
            if ($varName -eq 'SystemDrive') { [void]$roots.Add($val.TrimEnd('\', '/') + '\') }
        }
    }
    return @($roots | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}

function Get-EffectiveSpoolDirectory {
    <#
    .SYNOPSIS
        Resolves the real spool directory: parameter override -> registry -> OS default.
    .DESCRIPTION
        Honours HKLM\SYSTEM\CurrentControlSet\Control\Print\Printers\DefaultSpoolDirectory.
        The registry value is VALIDATED before it is trusted: a UNC path, a value
        shorter than 4 characters, or one containing control characters is almost
        always a misconfiguration, so it is ignored with a warning rather than
        being handed to the delete step (finding A6).
    .OUTPUTS
        PSCustomObject { Path, Source, RegistryValue, IsValid, Warning }
    #>
    [CmdletBinding()]
    param([AllowNull()][string]$Override)

    $systemRoot = $env:SystemRoot
    if ([string]::IsNullOrWhiteSpace($systemRoot)) { $systemRoot = 'C:\Windows' }
    $default = Join-Path -Path $systemRoot -ChildPath $Script:DefaultSpoolDirectory

    $result = [ordered]@{
        Path          = $default
        Source        = 'OSDefault'
        RegistryValue = $null
        IsValid       = $true
        Warning       = $null
    }

    if (-not [string]::IsNullOrWhiteSpace($Override)) {
        $result.Path = $Override
        $result.Source = 'Parameter'
        return [pscustomobject]$result
    }

    $regValue = $null
    try {
        if (Test-RunningOnWindows) {
            if (Test-Path -LiteralPath $Script:RegPrintersKey) {
                $item = Get-ItemProperty -LiteralPath $Script:RegPrintersKey -Name $Script:RegSpoolDirectoryVal -ErrorAction SilentlyContinue
                if ($item) { $regValue = [string]$item.($Script:RegSpoolDirectoryVal) }
            }
        }
    }
    catch {
        $result.Warning = "Registry probe failed: $($_.Exception.Message)"
    }
    $result.RegistryValue = $regValue

    if (-not [string]::IsNullOrWhiteSpace($regValue)) {
        $trimmed = $regValue.Trim().Trim('"')
        if ($trimmed -match '^[\\]{2}' -or $trimmed.Length -lt 4 -or $trimmed -match '[\x00-\x1f]') {
            $result.Warning = "Ignoring suspicious DefaultSpoolDirectory='$regValue'; using the OS default."
        }
        else {
            $result.Path = $trimmed
            $result.Source = 'Registry'
        }
    }
    return [pscustomobject]$result
}

function Test-PrintSpoolerAvailable {
    <#
    .SYNOPSIS
        True when the Print Spooler service exists on this host.
    .DESCRIPTION
        Server Core without the Print Services role, and Nano/IoT SKUs, have no
        Spooler. v4 reports ExitCode.UnsupportedPlatform (10) instead of a
        confusing "service not found" exception deep in the run.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    if (-not (Test-RunningOnWindows)) { return $false }
    try {
        $svc = Get-Service -Name $Script:SpoolerServiceName -ErrorAction SilentlyContinue
        return ($null -ne $svc)
    }
    catch {
        return $false
    }
}

function Get-WindowsSku {
    <#
    .SYNOPSIS
        Human-readable OS edition and build, used by diagnostics, logs and reports.
    #>
    [CmdletBinding()]
    param()

    $info = [ordered]@{ Caption = $null; Version = $null; Build = $null; Architecture = $null; ProductType = $null }
    try {
        $info.Architecture = $env:PROCESSOR_ARCHITECTURE
        if (Test-RunningOnWindows) {
            $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
            if ($os) {
                $info.Caption = $os.Caption
                $info.Version = [string]$os.Version
                $info.Build = [string]$os.BuildNumber
                $info.ProductType = switch ([int]$os.ProductType) {
                    1 { 'Workstation' }
                    2 { 'DomainController' }
                    3 { 'Server' }
                    default { "Unknown($($os.ProductType))" }
                }
            }
            else {
                $info.Version = [Environment]::OSVersion.Version.ToString()
                $info.Build = [Environment]::OSVersion.Version.Build.ToString()
            }
        }
        else {
            $info.Caption = [System.Runtime.InteropServices.RuntimeInformation]::OSDescription
            $info.Version = [Environment]::OSVersion.Version.ToString()
        }
    }
    catch {
        # An unqueryable OS still produces a usable (if sparse) record.
    }
    return [pscustomobject]$info
}

function Get-PrintSpoolerServiceDetail {
    <#
    .SYNOPSIS
        Rich service inventory: Spooler, its dependents and the per-user Print
        Workflow services, with start type and process id.
    #>
    [CmdletBinding()]
    param([string]$ComputerName = 'localhost')

    $sb = {
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($n in @('Spooler', 'PrintNotify', 'Fax')) {
            $svc = Get-Service -Name $n -ErrorAction SilentlyContinue
            if ($svc) {
                $wmi = Get-CimInstance -ClassName Win32_Service -Filter "Name='$n'" -ErrorAction SilentlyContinue
                $out.Add([pscustomobject]@{
                    Name        = $svc.Name
                    DisplayName = $svc.DisplayName
                    Status      = [string]$svc.Status
                    StartType   = if ($wmi) { $wmi.StartMode } else { [string]$svc.StartType }
                    PathName    = if ($wmi) { $wmi.PathName } else { '' }
                    ProcessId   = $svc.ServiceObjectId
                })
            }
        }
        # Per-user Print Workflow services are created dynamically; match by prefix.
        foreach ($svc in @(Get-Service -ErrorAction SilentlyContinue | Where-Object { $_.Name -like 'PrintWorkflowUserSvc*' })) {
            $out.Add([pscustomobject]@{
                Name        = $svc.Name
                DisplayName = $svc.DisplayName
                Status      = [string]$svc.Status
                StartType   = [string]$svc.StartType
                PathName    = ''
                ProcessId   = $svc.ServiceObjectId
            })
        }
        return $out.ToArray()
    }

    try {
        if (Test-IsLocalTarget -ComputerName $ComputerName) { return @(& $sb) }
        return @(Invoke-Command -ComputerName $ComputerName -ScriptBlock $sb -ErrorAction Stop)
    }
    catch {
        return @([pscustomobject]@{
                Name        = $Script:SpoolerServiceName
                DisplayName = 'n/a'
                Status      = 'Unreachable'
                StartType   = 'n/a'
                PathName    = ''
                ProcessId   = 0
                Error       = $_.Exception.Message
            })
    }
}

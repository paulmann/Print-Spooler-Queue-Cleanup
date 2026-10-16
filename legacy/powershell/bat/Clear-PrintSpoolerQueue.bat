@echo off
REM ===========================================================================
REM  Print Spooler Queue Cleanup - Windows batch launcher
REM ---------------------------------------------------------------------------
REM  Author  : Mikhail Deynekin
REM  E-Mail  : Mikhail@Deynekin.com
REM  Site    : https://Deynekin.com
REM  GitHub  : https://github.com/paulmann/Print-Spooler-Queue-Cleanup
REM  Version : @@SCRIPT_VERSION@@
REM  License : GPL-3.0-or-later
REM
REM  PURPOSE
REM    Thin, dependency-free front-end for Clear-PrintSpoolerQueue.ps1. It exists
REM    for the two situations where a .bat is genuinely the right answer:
REM    double-clicking from Explorer, and handing a technician a file that runs
REM    in a bare cmd prompt. Everything substantive happens in PowerShell; this
REM    file only resolves an engine, elevates, applies a staged update,
REM    translates arguments and propagates the exit code.
REM
REM  WHY IT IS THIN (lesson from v3)
REM    v3.1.0 reimplemented the whole cleanup in batch. Two implementations of
REM    one destructive operation is two chances to be wrong, and the batch copy
REM    could never offer remote, selective or backed-up cleanup. v4 delegates, so
REM    there is exactly one place where a file can be deleted.
REM
REM  ENGINE SELECTION
REM    1. pwsh.exe       (PowerShell 7.x) - required for -ThrottleLimit > 1
REM    2. powershell.exe (Windows PowerShell 5.1) - present on every Windows
REM    Override with /PS7, /PS5, or the PSQ_ENGINE environment variable.
REM
REM  EXIT CODES (identical to the PowerShell script, so monitoring rules port)
REM    0   Success                        8   Update applied
REM    1   Partial success                9   Update failed
REM    2   Failed                        10   Unsupported platform
REM    3   Bad arguments                 11   Diagnostics only (informational)
REM    4   Cancelled by user             12   Spooler service failure
REM    5   Not administrator             13   Restore failed
REM    6   Update check failed           14   Signature invalid
REM    7   Update available              15   Watchdog timeout
REM ===========================================================================

setlocal EnableExtensions EnableDelayedExpansion
title Print Spooler Queue Cleanup v@@SCRIPT_VERSION@@

REM Captured BEFORE the parse loop consumes the arguments with `shift`, so that
REM /Elevate can re-launch with exactly what the user typed.
set "ORIG_ARGS=%*"

set "SCRIPT_VERSION=@@SCRIPT_VERSION@@"
set "SCRIPT_NAME=%~nx0"
set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%Clear-PrintSpoolerQueue.ps1"
set "PRODUCT_FAMILY=PrintSpoolerCleanup"
set "REPO_URL=https://github.com/paulmann/Print-Spooler-Queue-Cleanup"
set "STAGE_DIR=%SCRIPT_DIR%.update"

REM --- parsed flags -----------------------------------------------------------
set "FLAG_FORCE=0"
set "FLAG_ELEVATE=0"
set "FLAG_QUIET=0"
set "FLAG_NOPAUSE=0"
set "FLAG_HELP=0"
set "FLAG_VERSION=0"
set "FLAG_DIAGNOSTIC=0"
set "FLAG_PS7=0"
set "FLAG_PS5=0"
set "ARG_LOGPATH="
set "ARG_COMPUTERS="
set "ARG_HELPSECTION="
set "ENGINE="
set "EXIT_CODE=0"
set "PS_ARGS="

REM ---------------------------------------------------------------------------
REM  Interactive detection
REM  A "press any key" pause under Task Scheduler blocks the task until its
REM  execution time limit kills it - exactly how the v3 README example failed.
REM  cmd.exe sets CMDCMDLINE differently for each case that matters:
REM    double-clicked  -> cmd /c ""C:\path\file.bat" "   -> pause
REM    typed in a console -> "C:\path\file.bat" /Force   -> pause
REM    Task Scheduler  -> cmd /c C:\path\file.bat -Force  -> never pause
REM  /NoPause always wins. The PowerShell side performs its own, stronger check
REM  ([Console]::IsInputRedirected) and refuses to prompt regardless.
REM ---------------------------------------------------------------------------
set "FLAG_INTERACTIVE=1"
echo %CMDCMDLINE% | find /I "/c %~f0" >nul 2>&1 && set "FLAG_INTERACTIVE=0"
echo %CMDCMDLINE% | find /I "/c ""%~f0" >nul 2>&1 && set "FLAG_INTERACTIVE=1"
if defined PSQ_NONINTERACTIVE set "FLAG_INTERACTIVE=0"

REM ---------------------------------------------------------------------------
REM  Argument parsing
REM ---------------------------------------------------------------------------
:parse_args
if "%~1"=="" goto args_done
set "ARG=%~1"

REM --- a switch that takes a value MUST be given one ------------------------
REM An empty value is not harmless here. PowerShell treats '' as FALSY, so the
REM script's selection guard `if ($Options['PrinterName'] -and ...)` does not fire
REM and a blank -PrinterName is IGNORED rather than matched: the run silently
REM widens to EVERY printer and every job. A blank -ExcludePrinterName is worse -
REM it withdraws a protection the operator asked for. Typing `/printer` and
REM forgetting the value is an easy mistake, so it is refused here with exit
REM code 3 instead of being forwarded.
set "NEEDS_VALUE="
for %%S in (/log -log /logpath /computers -computers /computername /hosts -hosts /printer -printer /excludeprinter -excludeprinter /olderthan -olderthan /newerthan /jobid /owner /backuppath /archivepath /exportpath /exportformat /restore /restoredir /config /saveconfig /throttle /timeout /settle /loglevel /retentiondays /channel /bundle) do (
    if /I "%ARG%"=="%%S" set "NEEDS_VALUE=1"
)
if defined NEEDS_VALUE if "%~2"=="" (
    echo ERROR: %ARG% requires a value.
    echo Run "%SCRIPT_NAME% /?" for usage.
    set "EXIT_CODE=3"
    goto finish
)

if /I "%ARG%"=="/?"                set "FLAG_HELP=1" & shift & goto parse_args
if /I "%ARG%"=="-?"                set "FLAG_HELP=1" & shift & goto parse_args
if /I "%ARG%"=="/h"                set "FLAG_HELP=1" & shift & goto parse_args
if /I "%ARG%"=="-h"                set "FLAG_HELP=1" & shift & goto parse_args
if /I "%ARG%"=="/help"             if "%~2"=="" (set "FLAG_HELP=1" & shift & goto parse_args) else (set "FLAG_HELP=1" & set "ARG_HELPSECTION=%~2" & shift & shift & goto parse_args)
if /I "%ARG%"=="--help"            set "FLAG_HELP=1" & shift & goto parse_args
if /I "%ARG%"=="/version"          set "FLAG_VERSION=1" & shift & goto parse_args
if /I "%ARG%"=="--version"         set "FLAG_VERSION=1" & shift & goto parse_args
if /I "%ARG%"=="-v"                set "FLAG_VERSION=1" & shift & goto parse_args

if /I "%ARG%"=="/force"            set "FLAG_FORCE=1" & shift & goto parse_args
if /I "%ARG%"=="-force"            set "FLAG_FORCE=1" & shift & goto parse_args
if /I "%ARG%"=="/y"                set "FLAG_FORCE=1" & shift & goto parse_args
if /I "%ARG%"=="-y"                set "FLAG_FORCE=1" & shift & goto parse_args
if /I "%ARG%"=="/elevate"          set "FLAG_ELEVATE=1" & shift & goto parse_args
if /I "%ARG%"=="-elevate"          set "FLAG_ELEVATE=1" & shift & goto parse_args
if /I "%ARG%"=="/quiet"            set "FLAG_QUIET=1" & shift & goto parse_args
if /I "%ARG%"=="-quiet"            set "FLAG_QUIET=1" & shift & goto parse_args
if /I "%ARG%"=="/nopause"          set "FLAG_NOPAUSE=1" & set "FLAG_INTERACTIVE=0" & shift & goto parse_args
if /I "%ARG%"=="-nopause"          set "FLAG_NOPAUSE=1" & set "FLAG_INTERACTIVE=0" & shift & goto parse_args
if /I "%ARG%"=="/ps7"              set "FLAG_PS7=1" & shift & goto parse_args
if /I "%ARG%"=="-ps7"              set "FLAG_PS7=1" & shift & goto parse_args
if /I "%ARG%"=="/ps5"              set "FLAG_PS5=1" & shift & goto parse_args
if /I "%ARG%"=="-ps5"              set "FLAG_PS5=1" & shift & goto parse_args
if /I "%ARG%"=="/diagnostic"       set "FLAG_DIAGNOSTIC=1" & shift & goto parse_args
if /I "%ARG%"=="-diagnostic"       set "FLAG_DIAGNOSTIC=1" & shift & goto parse_args

REM --- switches that map 1:1 onto the PowerShell script ----------------------
if /I "%ARG%"=="/whatif"           set "PS_ARGS=!PS_ARGS! -WhatIf" & shift & goto parse_args
if /I "%ARG%"=="-whatif"           set "PS_ARGS=!PS_ARGS! -WhatIf" & shift & goto parse_args
if /I "%ARG%"=="/dry-run"          set "PS_ARGS=!PS_ARGS! -WhatIf" & shift & goto parse_args
if /I "%ARG%"=="/repair"           set "PS_ARGS=!PS_ARGS! -Repair" & shift & goto parse_args
if /I "%ARG%"=="-repair"           set "PS_ARGS=!PS_ARGS! -Repair" & shift & goto parse_args
if /I "%ARG%"=="/backup"           set "PS_ARGS=!PS_ARGS! -Backup" & shift & goto parse_args
if /I "%ARG%"=="-backup"           set "PS_ARGS=!PS_ARGS! -Backup" & shift & goto parse_args
if /I "%ARG%"=="/checkupdate"      set "PS_ARGS=!PS_ARGS! -CheckUpdate" & shift & goto parse_args
if /I "%ARG%"=="-checkupdate"      set "PS_ARGS=!PS_ARGS! -CheckUpdate" & shift & goto parse_args
if /I "%ARG%"=="/update"           set "PS_ARGS=!PS_ARGS! -Update -Force" & shift & goto parse_args
if /I "%ARG%"=="-update"           set "PS_ARGS=!PS_ARGS! -Update -Force" & shift & goto parse_args
if /I "%ARG%"=="/autoupdate"       set "PS_ARGS=!PS_ARGS! -AutoUpdate" & shift & goto parse_args
if /I "%ARG%"=="-autoupdate"       set "PS_ARGS=!PS_ARGS! -AutoUpdate" & shift & goto parse_args
if /I "%ARG%"=="/selftest"         set "PS_ARGS=!PS_ARGS! -SelfTest" & shift & goto parse_args
if /I "%ARG%"=="-selftest"         set "PS_ARGS=!PS_ARGS! -SelfTest" & shift & goto parse_args
if /I "%ARG%"=="/watchdog"         set "PS_ARGS=!PS_ARGS! -Watchdog" & shift & goto parse_args
if /I "%ARG%"=="-watchdog"         set "PS_ARGS=!PS_ARGS! -Watchdog" & shift & goto parse_args
if /I "%ARG%"=="/eventlog"         set "PS_ARGS=!PS_ARGS! -EventLog" & shift & goto parse_args
if /I "%ARG%"=="-eventlog"         set "PS_ARGS=!PS_ARGS! -EventLog" & shift & goto parse_args
if /I "%ARG%"=="/jsonlog"          set "PS_ARGS=!PS_ARGS! -JsonLog" & shift & goto parse_args
if /I "%ARG%"=="-jsonlog"          set "PS_ARGS=!PS_ARGS! -JsonLog" & shift & goto parse_args
if /I "%ARG%"=="/json"             set "PS_ARGS=!PS_ARGS! -MachineReadableJson -Quiet -NoProgress" & shift & goto parse_args
if /I "%ARG%"=="-json"             set "PS_ARGS=!PS_ARGS! -MachineReadableJson -Quiet -NoProgress" & shift & goto parse_args
if /I "%ARG%"=="/ascii"            set "PS_ARGS=!PS_ARGS! -Ascii" & shift & goto parse_args
if /I "%ARG%"=="-ascii"            set "PS_ARGS=!PS_ARGS! -Ascii" & shift & goto parse_args
if /I "%ARG%"=="/nocolor"          set "PS_ARGS=!PS_ARGS! -NoColor" & shift & goto parse_args
if /I "%ARG%"=="-nocolor"          set "PS_ARGS=!PS_ARGS! -NoColor" & shift & goto parse_args
if /I "%ARG%"=="/purgeforms"       set "PS_ARGS=!PS_ARGS! -PurgeFormsCache" & shift & goto parse_args
if /I "%ARG%"=="-purgeforms"       set "PS_ARGS=!PS_ARGS! -PurgeFormsCache" & shift & goto parse_args
if /I "%ARG%"=="/purgedrivers"     set "PS_ARGS=!PS_ARGS! -PurgeDriverStore" & shift & goto parse_args
if /I "%ARG%"=="-purgedrivers"     set "PS_ARGS=!PS_ARGS! -PurgeDriverStore" & shift & goto parse_args
if /I "%ARG%"=="/resetspool"       set "PS_ARGS=!PS_ARGS! -ResetSpoolDirectory" & shift & goto parse_args
if /I "%ARG%"=="/verifysignature"  set "PS_ARGS=!PS_ARGS! -VerifySignature" & shift & goto parse_args
if /I "%ARG%"=="-verifysignature"  set "PS_ARGS=!PS_ARGS! -VerifySignature" & shift & goto parse_args
if /I "%ARG%"=="/noninteractive"   set "PS_ARGS=!PS_ARGS! -NonInteractive" & shift & goto parse_args
if /I "%ARG%"=="-noninteractive"   set "PS_ARGS=!PS_ARGS! -NonInteractive" & shift & goto parse_args
if /I "%ARG%"=="/selftestprint"    set "PS_ARGS=!PS_ARGS! -SelfTestPrint" & shift & goto parse_args
if /I "%ARG%"=="/bundle"           set "PS_ARGS=!PS_ARGS! -DiagnosticsBundlePath %~2" & shift & shift & goto parse_args

REM --- switches that take a value --------------------------------------------
if /I "%ARG%"=="/log"              set "ARG_LOGPATH=%~2" & shift & shift & goto parse_args
if /I "%ARG%"=="-log"              set "ARG_LOGPATH=%~2" & shift & shift & goto parse_args
if /I "%ARG%"=="/logpath"          set "ARG_LOGPATH=%~2" & shift & shift & goto parse_args
if /I "%ARG%"=="/computers"        set "ARG_COMPUTERS=%~2" & shift & shift & goto parse_args
if /I "%ARG%"=="-computers"        set "ARG_COMPUTERS=%~2" & shift & shift & goto parse_args
if /I "%ARG%"=="/computername"     set "ARG_COMPUTERS=%~2" & shift & shift & goto parse_args
if /I "%ARG%"=="/hosts"            set "PS_ARGS=!PS_ARGS! -ComputerListFile "%~2"" & shift & shift & goto parse_args
if /I "%ARG%"=="-hosts"            set "PS_ARGS=!PS_ARGS! -ComputerListFile "%~2"" & shift & shift & goto parse_args
if /I "%ARG%"=="/printer"          set "PS_ARGS=!PS_ARGS! -PrinterName "%~2"" & shift & shift & goto parse_args
if /I "%ARG%"=="-printer"          set "PS_ARGS=!PS_ARGS! -PrinterName "%~2"" & shift & shift & goto parse_args
if /I "%ARG%"=="/excludeprinter"   set "PS_ARGS=!PS_ARGS! -ExcludePrinterName "%~2"" & shift & shift & goto parse_args
if /I "%ARG%"=="-excludeprinter"   set "PS_ARGS=!PS_ARGS! -ExcludePrinterName "%~2"" & shift & shift & goto parse_args
if /I "%ARG%"=="/olderthan"        set "PS_ARGS=!PS_ARGS! -OlderThan %~2" & shift & shift & goto parse_args
if /I "%ARG%"=="-olderthan"        set "PS_ARGS=!PS_ARGS! -OlderThan %~2" & shift & shift & goto parse_args
if /I "%ARG%"=="/newerthan"        set "PS_ARGS=!PS_ARGS! -NewerThan %~2" & shift & shift & goto parse_args
if /I "%ARG%"=="/jobid"            set "PS_ARGS=!PS_ARGS! -JobId %~2" & shift & shift & goto parse_args
if /I "%ARG%"=="/owner"            set "PS_ARGS=!PS_ARGS! -OwnerName "%~2"" & shift & shift & goto parse_args
if /I "%ARG%"=="/backuppath"       set "PS_ARGS=!PS_ARGS! -BackupPath "%~2"" & shift & shift & goto parse_args
if /I "%ARG%"=="/archivepath"      set "PS_ARGS=!PS_ARGS! -ArchivePath "%~2"" & shift & shift & goto parse_args
if /I "%ARG%"=="/exportpath"       set "PS_ARGS=!PS_ARGS! -ExportPath "%~2"" & shift & shift & goto parse_args
if /I "%ARG%"=="/exportformat"     set "PS_ARGS=!PS_ARGS! -ExportFormat %~2" & shift & shift & goto parse_args
if /I "%ARG%"=="/restore"          set "PS_ARGS=!PS_ARGS! -Restore "%~2"" & shift & shift & goto parse_args
if /I "%ARG%"=="/restoredir"       set "PS_ARGS=!PS_ARGS! -RestoreDestination "%~2"" & shift & shift & goto parse_args
if /I "%ARG%"=="/config"           set "PS_ARGS=!PS_ARGS! -ConfigPath "%~2"" & shift & shift & goto parse_args
if /I "%ARG%"=="/saveconfig"       set "PS_ARGS=!PS_ARGS! -SaveConfig -ConfigPath "%~2"" & shift & shift & goto parse_args
if /I "%ARG%"=="/throttle"         set "PS_ARGS=!PS_ARGS! -ThrottleLimit %~2" & shift & shift & goto parse_args
if /I "%ARG%"=="/timeout"          set "PS_ARGS=!PS_ARGS! -ServiceTimeoutSeconds %~2" & shift & shift & goto parse_args
if /I "%ARG%"=="/settle"           set "PS_ARGS=!PS_ARGS! -SettleSeconds %~2" & shift & shift & goto parse_args
if /I "%ARG%"=="/loglevel"         set "PS_ARGS=!PS_ARGS! -LogLevel %~2" & shift & shift & goto parse_args
if /I "%ARG%"=="/retentiondays"    set "PS_ARGS=!PS_ARGS! -LogRetentionDays %~2" & shift & shift & goto parse_args
if /I "%ARG%"=="/channel"          set "PS_ARGS=!PS_ARGS! -UpdateChannel %~2" & shift & shift & goto parse_args

echo ERROR: Unknown argument: %ARG%
echo Run "%SCRIPT_NAME% /?" for usage.
set "EXIT_CODE=3"
goto finish

:args_done

REM ---------------------------------------------------------------------------
REM  Help and version are served locally when they need no engine, and by
REM  PowerShell when a specific reference section was requested.
REM ---------------------------------------------------------------------------
if "%FLAG_VERSION%"=="1" goto show_version
if "%FLAG_HELP%"=="1" (
    if not "%ARG_HELPSECTION%"=="" goto run_powershell
    call :show_help
    set "EXIT_CODE=0"
    goto finish
)

REM ---------------------------------------------------------------------------
REM  Locate the PowerShell script
REM ---------------------------------------------------------------------------
if not exist "%PS_SCRIPT%" (
    call :msg ERROR "Clear-PrintSpoolerQueue.ps1 was not found next to this launcher:"
    call :msg ERROR "    %PS_SCRIPT%"
    call :msg INFO  "Download both files from %REPO_URL%"
    set "EXIT_CODE=2"
    goto finish
)

REM ---------------------------------------------------------------------------
REM  Choose the PowerShell engine
REM ---------------------------------------------------------------------------
call :resolve_engine
if not defined ENGINE (
    call :msg ERROR "No PowerShell engine was found. Windows PowerShell 5.1 or later is required."
    set "EXIT_CODE=10"
    goto finish
)
if "%FLAG_QUIET%"=="0" call :msg INFO "Engine: !ENGINE!"

REM ---------------------------------------------------------------------------
REM  Apply a staged update BEFORE running, so this run uses the new version.
REM ---------------------------------------------------------------------------
call :apply_pending_update

REM ---------------------------------------------------------------------------
REM  Elevation
REM ---------------------------------------------------------------------------
if "%FLAG_HELP%"=="1" goto run_powershell
if "%FLAG_VERSION%"=="1" goto run_powershell

call :check_admin
if errorlevel 1 (
    if "%FLAG_ELEVATE%"=="1" (
        call :elevate
        goto finish
    )
    call :msg ERROR "Administrator privileges are required to manage the Print Spooler."
    call :msg INFO  "Right-click and choose 'Run as administrator', or re-run with /Elevate."
    set "EXIT_CODE=5"
    goto finish
)

REM ---------------------------------------------------------------------------
REM  Run
REM ---------------------------------------------------------------------------
:run_powershell
if "%FLAG_FORCE%"=="1"       set "PS_ARGS=!PS_ARGS! -Force"
if "%FLAG_DIAGNOSTIC%"=="1"  set "PS_ARGS=!PS_ARGS! -Diagnostic"
if "%FLAG_QUIET%"=="1"       set "PS_ARGS=!PS_ARGS! -Quiet -NoProgress"
if "%FLAG_INTERACTIVE%"=="0" set "PS_ARGS=!PS_ARGS! -NonInteractive -NoProgress"
if not "%ARG_HELPSECTION%"=="" set "PS_ARGS=!PS_ARGS! -HelpSection %ARG_HELPSECTION%"
if defined ARG_LOGPATH       set "PS_ARGS=!PS_ARGS! -LogPath "%ARG_LOGPATH%""
if defined ARG_COMPUTERS     set "PS_ARGS=!PS_ARGS! -ComputerName %ARG_COMPUTERS%"

if "%FLAG_QUIET%"=="0" (
    echo ============================================================================
    echo     Print Spooler Queue Cleanup v%SCRIPT_VERSION%
    echo     Author : Mikhail Deynekin ^<Mikhail@Deynekin.com^>  https://Deynekin.com
    echo     GitHub : %REPO_URL%
    echo ============================================================================
    echo.
    call :msg INFO "Running: !ENGINE! -File "%PS_SCRIPT%"!PS_ARGS!"
    echo.
)

"!ENGINE!" -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%PS_SCRIPT%" !PS_ARGS!
set "EXIT_CODE=!ERRORLEVEL!"
goto finish

REM ---------------------------------------------------------------------------
:finish
if "%FLAG_INTERACTIVE%"=="1" if "%FLAG_QUIET%"=="0" (
    echo.
    echo Exit code: %EXIT_CODE%
    echo Press any key to close . . .
    pause >nul
)
endlocal & exit /b %EXIT_CODE%


REM ===========================================================================
REM  Subroutines
REM ===========================================================================

REM ---------------------------------------------------------------------------
REM  :msg <LEVEL> <TEXT>
REM ---------------------------------------------------------------------------
:msg
    if "%FLAG_QUIET%"=="1" goto :eof
    echo [%~1] %~2
    goto :eof

REM ---------------------------------------------------------------------------
REM  :resolve_engine  ->  sets ENGINE to a full path, or leaves it undefined
REM ---------------------------------------------------------------------------
:resolve_engine
    if defined PSQ_ENGINE (
        if /I "%PSQ_ENGINE%"=="pwsh" set "FLAG_PS7=1"
        if /I "%PSQ_ENGINE%"=="powershell" set "FLAG_PS5=1"
    )
    set "PSEXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

    if "%FLAG_PS5%"=="1" (
        if exist "%PSEXE%" ( set "ENGINE=%PSEXE%" ) else ( call :msg ERROR "/PS5 was requested but powershell.exe was not found." )
        goto :eof
    )

    rem PowerShell 7 first: it is the only engine with ForEach-Object -Parallel,
    rem which is what makes fleet-wide cleanup fast instead of serial.
    where pwsh.exe >nul 2>&1
    if not errorlevel 1 (
        for /f "usebackq delims=" %%P in (`where pwsh.exe 2^>nul`) do (
            if not defined ENGINE if exist "%%P" set "ENGINE=%%P"
        )
        if defined ENGINE goto :eof
    )

    if "%FLAG_PS7%"=="1" (
        call :msg ERROR "/PS7 was requested but pwsh.exe was not found in PATH."
        goto :eof
    )

    if exist "%PSEXE%" ( set "ENGINE=%PSEXE%" & goto :eof )
    goto :eof

REM ---------------------------------------------------------------------------
REM  :check_admin   errorlevel 0 = elevated, 1 = not elevated
REM    Three independent probes: `net session` fails on hosts where the Server
REM    service is stopped even though the token IS elevated; `fltmc` and
REM    `openfiles` cover the remaining cases.
REM ---------------------------------------------------------------------------
:check_admin
    net session >nul 2>&1
    if not errorlevel 1 exit /b 0
    fltmc >nul 2>&1
    if not errorlevel 1 exit /b 0
    openfiles >nul 2>&1
    if not errorlevel 1 exit /b 0
    exit /b 1

REM ---------------------------------------------------------------------------
REM  :elevate
REM    Re-launches this launcher elevated via UAC, forwarding the ORIGINAL
REM    arguments (%ORIG_ARGS%) because the parse loop has already shifted them
REM    away. /NoPause is appended so the elevated window behaves predictably.
REM ---------------------------------------------------------------------------
:elevate
    call :msg INFO "Requesting administrator privileges (UAC prompt)..."
    set "CHILD_ARGS=%ORIG_ARGS% /NoPause"
    set "VBS_FILE=%TEMP%\psq-elevate-%RANDOM%%RANDOM%.vbs"
    > "!VBS_FILE!" echo Set sh = CreateObject("Shell.Application")
    >>"!VBS_FILE!" echo sh.ShellExecute "%~f0", "!CHILD_ARGS!", "%SCRIPT_DIR%", "runas", 1
    cscript //nologo "!VBS_FILE!" >nul 2>&1
    set "RC=!ERRORLEVEL!"
    del "!VBS_FILE!" >nul 2>&1
    if not "!RC!"=="0" (
        call :msg ERROR "Elevation was declined or failed (code !RC!)."
        set "EXIT_CODE=5"
    ) else (
        call :msg OK "Elevated instance launched. This window can be closed."
        set "EXIT_CODE=0"
    )
    goto :eof

REM ---------------------------------------------------------------------------
REM  :apply_pending_update
REM    Moves files staged by a previous /Update into place.
REM
REM    The SHA-256 of every staged file was verified twice already: once when it
REM    was downloaded (against the release's checksums.txt) and once when it was
REM    written into pending-update.json. This swap therefore only moves
REM    already-verified files, which keeps the launcher free of both a hashing
REM    dependency and the PowerShell-in-batch quoting hazards that make such
REM    one-liners fragile. The PowerShell entry point performs the same swap with
REM    full hash re-verification for anyone who runs the .ps1 directly.
REM ---------------------------------------------------------------------------
:apply_pending_update
    if not exist "!STAGE_DIR!\pending-update.json" goto :eof
    if "%FLAG_QUIET%"=="0" call :msg INFO "A staged update was found; applying it before this run..."

    if exist "!STAGE_DIR!\Clear-PrintSpoolerQueue.ps1" (
        move /Y "!STAGE_DIR!\Clear-PrintSpoolerQueue.ps1" "%PS_SCRIPT%" >nul 2>&1
        if errorlevel 1 (
            call :msg WARN "Could not replace Clear-PrintSpoolerQueue.ps1 (in use?). The update stays staged."
            goto :eof
        )
        if "%FLAG_QUIET%"=="0" call :msg OK "Clear-PrintSpoolerQueue.ps1 updated."
    )

    rem A running .bat cannot reliably replace itself - cmd.exe re-reads the file
    rem by byte offset while it executes. Skipping it here is deliberate; the
    rem swap happens on the next start.
    if exist "!STAGE_DIR!\Clear-PrintSpoolerQueue.bat" (
        if "%FLAG_QUIET%"=="0" call :msg INFO "The launcher update will be applied on the next start."
    )
    goto :eof

REM ---------------------------------------------------------------------------
REM  :show_help
REM ---------------------------------------------------------------------------
:show_help
    echo.
    echo Print Spooler Queue Cleanup v%SCRIPT_VERSION% - launcher
    echo Author: Mikhail Deynekin ^<Mikhail@Deynekin.com^>   https://Deynekin.com
    echo GitHub: %REPO_URL%
    echo.
    echo USAGE:
    echo     %SCRIPT_NAME% [options]
    echo     %SCRIPT_NAME% /?                    this help
    echo     %SCRIPT_NAME% /Help ^<section^>       a section of the full reference
    echo     %SCRIPT_NAME% /Version              version information
    echo.
    echo MODES:
    echo     /Diagnostic        Read-only health and queue inspection. Changes nothing.
    echo     /WhatIf            Dry run: full inventory and plan, no changes.
    echo     /Repair            Deep clean: also cycles dependent services and can
    echo                        purge the Forms cache or reset a bad spool directory.
    echo     /Restore ^<zip^>      Restore a backup archive created by /Backup.
    echo     /CheckUpdate       Ask GitHub whether a newer version exists, then exit.
    echo     /Update            Download, verify (SHA-256) and install a newer version.
    echo     /SelfTest          Verify that this copy of the tool is internally sound.
    echo.
    echo SELECTION (default: every queued job):
    echo     /Printer ^<name^>     Only these printers. Wildcards allowed: "HP*"
    echo     /ExcludePrinter ^<n^> Never touch these printers.
    echo     /OlderThan ^<age^>    Only jobs older than this: 30s 15m 2h 7d
    echo     /NewerThan ^<age^>    Only jobs newer than this.
    echo     /JobId ^<ids^>        Only these job ids, e.g. 4217,4218
    echo     /Owner ^<name^>      Only jobs submitted by this user.
    echo     /Computers ^<list^>   Comma-separated remote targets (needs WinRM).
    echo     /Hosts ^<file^>       Text file with one computer name per line.
    echo     /Throttle ^<n^>       Remote workers in parallel (PowerShell 7 only).
    echo.
    echo SAFETY:
    echo     /Backup            Archive every file before deleting it. STRONGLY
    echo                        recommended. Default location:
    echo                        %%ProgramData%%\%PRODUCT_FAMILY%\Backups
    echo     /BackupPath ^<dir^>   Where to put the backup archives.
    echo     /ArchivePath ^<dir^> Move the archives here afterwards (e.g. a log share).
    echo     /Force  or  /Y     Skip the confirmation prompt (required unattended).
    echo     /Elevate           Request UAC elevation instead of failing.
    echo     /NonInteractive    Never prompt; fail with exit code 4 instead.
    echo     /VerifySignature   Verify the Authenticode signature before running.
    echo     /Watchdog          Terminate and restart the Spooler if the run hangs.
    echo     /Timeout ^<sec^>     Service stop/start timeout (default 60).
    echo     /SelfTestPrint     After cleanup, submit and cancel a test job to prove
    echo                        that printing really works again.
    echo.
    echo REPAIR OPTIONS (require /Repair):
    echo     /PurgeForms        Clear the spooler Forms cache.
    echo     /PurgeDrivers      Remove driver packages no printer references (needs /Force).
    echo     /ResetSpool        Remove a bad DefaultSpoolDirectory registry override.
    echo.
    echo OUTPUT:
    echo     /Quiet             No decorative output.
    echo     /NoPause           Do not wait for a keypress before closing.
    echo     /NoColor           Plain monochrome output.
    echo     /Ascii             ASCII glyphs for legacy code pages.
    echo     /Json              One JSON document on stdout, nothing else (for RMM).
    echo     /JsonLog           Write the log file as JSON Lines.
    echo     /Log ^<file^>         Explicit log file path.
    echo     /LogLevel ^<lvl^>    TRACE DEBUG INFO WARN ERROR CRITICAL
    echo     /RetentionDays ^<n^> Delete rotated logs older than n days.
    echo     /ExportFormat ^<f^>  Json, Csv, Html or Text report.
    echo     /ExportPath ^<dir^>  Where to write the report.
    echo     /Bundle ^<dir^>      Write a diagnostics ZIP suitable for a bug report.
    echo     /EventLog          Also write to the Windows Event Log.
    echo.
    echo ENGINE:
    echo     /PS7               Require PowerShell 7 (pwsh.exe).
    echo     /PS5               Require Windows PowerShell 5.1.
    echo     (default: pwsh.exe when present, otherwise powershell.exe)
    echo.
    echo EXIT CODES:
    echo      0 Success                   8 Update applied
    echo      1 Partial success           9 Update failed
    echo      2 Failed                   10 Unsupported platform
    echo      3 Bad arguments            11 Diagnostics only (not an error)
    echo      4 Cancelled by user        12 Spooler service failure
    echo      5 Not administrator        13 Restore failed
    echo      6 Update check failed      14 Signature invalid
    echo      7 Update available         15 Watchdog timeout
    echo.
    echo EXAMPLES:
    echo     %SCRIPT_NAME% /Diagnostic
    echo     %SCRIPT_NAME% /Force /Backup
    echo     %SCRIPT_NAME% /Force /Backup /OlderThan 2h
    echo     %SCRIPT_NAME% /Repair /Force /Backup /PurgeForms
    echo     %SCRIPT_NAME% /CheckUpdate
    echo     %SCRIPT_NAME% /Computers SRV01,SRV02 /Force /Backup /Throttle 8 /PS7
    echo     %SCRIPT_NAME% /Force /Quiet /NoPause /Backup /OlderThan 6h    ^(scheduled^)
    echo     %SCRIPT_NAME% /Help Safety
    echo.
    goto :eof

REM ---------------------------------------------------------------------------
REM  :show_version
REM ---------------------------------------------------------------------------
:show_version
    echo %SCRIPT_NAME% v%SCRIPT_VERSION%
    echo %REPO_URL%
    if exist "%PS_SCRIPT%" (
        for /f "usebackq delims=" %%V in (`findstr /R /C:"ScriptVersion       = " "%PS_SCRIPT%" 2^>nul`) do echo %%V
    )
    set "EXIT_CODE=0"
    goto finish

@echo off
REM ============================================================================
REM  Print Spooler Queue Cleanup (Windows batch front-end)
REM ----------------------------------------------------------------------------
REM  Author : Mikhail Deynekin
REM  E-Mail : Mikhail@Deynekin.com
REM  Site   : https://Deynekin.com
REM  GitHub : https://github.com/paulmann/Print-Spooler-Queue-Cleanup
REM
REM  Version: 3.1.0 (full rewrite to match the PowerShell script's safety model:
REM           robust admin check, locale-independent size accounting via
REM           PowerShell helper, predictable exit codes, /WhatIf and /Force,
REM           best-effort spooler recovery on failure, and an opt-in /UsePS
REM           mode that delegates the cleanup to Clear-PrintSpoolerQueue.ps1).
REM
REM  Description:
REM    Stops the Windows Print Spooler service, removes residual *.SHD/*.SPL
REM    files from the active spool directory, and restarts the service.
REM    The deletion step is intentionally limited to *.SHD and *.SPL — no other
REM    files in the spool directory are ever touched.
REM
REM  Exit codes:
REM    0  Success
REM    1  Generic / fatal error (admin missing, spool dir missing, service
REM       failed to stop or start, etc.)
REM    2  Completed with warnings (some files could not be deleted, but the
REM       Spooler is running again)
REM    3  Bad command line arguments
REM    4  Cancelled by user (interactive confirmation declined)
REM ============================================================================

setlocal EnableExtensions EnableDelayedExpansion

REM --- Constants ---------------------------------------------------------------
set "SCRIPT_VERSION=3.1.0"
set "SCRIPT_NAME=%~nx0"
set "SCRIPT_DIR=%~dp0"
set "SERVICE_NAME=Spooler"
set "PS_SCRIPT=%SCRIPT_DIR%Clear-PrintSpoolerQueue.ps1"

REM --- Defaults / parsed flags -------------------------------------------------
set "FLAG_FORCE=0"
set "FLAG_WHATIF=0"
set "FLAG_USEPS=0"
set "FLAG_QUIET=0"
set "FLAG_NOPAUSE=0"
set "ARG_LOGPATH="

REM --- Argument parsing --------------------------------------------------------
:parse_args
if "%~1"=="" goto args_done
set "_arg=%~1"
if /I "!_arg!"=="/?"        goto show_help
if /I "!_arg!"=="-?"        goto show_help
if /I "!_arg!"=="/h"        goto show_help
if /I "!_arg!"=="-h"        goto show_help
if /I "!_arg!"=="/help"     goto show_help
if /I "!_arg!"=="--help"    goto show_help
if /I "!_arg!"=="/version"  goto show_version
if /I "!_arg!"=="--version" goto show_version
if /I "!_arg!"=="-v"        goto show_version
if /I "!_arg!"=="/force"    ( set "FLAG_FORCE=1"   & shift & goto parse_args )
if /I "!_arg!"=="-force"    ( set "FLAG_FORCE=1"   & shift & goto parse_args )
if /I "!_arg!"=="/y"        ( set "FLAG_FORCE=1"   & shift & goto parse_args )
if /I "!_arg!"=="/whatif"   ( set "FLAG_WHATIF=1"  & shift & goto parse_args )
if /I "!_arg!"=="-whatif"   ( set "FLAG_WHATIF=1"  & shift & goto parse_args )
if /I "!_arg!"=="/dry-run"  ( set "FLAG_WHATIF=1"  & shift & goto parse_args )
if /I "!_arg!"=="/useps"    ( set "FLAG_USEPS=1"   & shift & goto parse_args )
if /I "!_arg!"=="-useps"    ( set "FLAG_USEPS=1"   & shift & goto parse_args )
if /I "!_arg!"=="/quiet"    ( set "FLAG_QUIET=1"   & shift & goto parse_args )
if /I "!_arg!"=="-quiet"    ( set "FLAG_QUIET=1"   & shift & goto parse_args )
if /I "!_arg!"=="/nopause"  ( set "FLAG_NOPAUSE=1" & shift & goto parse_args )
if /I "!_arg!"=="-nopause"  ( set "FLAG_NOPAUSE=1" & shift & goto parse_args )
if /I "!_arg!"=="/log"      ( set "ARG_LOGPATH=%~2" & shift & shift & goto parse_args )
if /I "!_arg!"=="-log"      ( set "ARG_LOGPATH=%~2" & shift & shift & goto parse_args )
echo ERROR: Unknown argument: !_arg!
echo Run "%SCRIPT_NAME% /?" for usage.
endlocal & exit /b 3
:args_done

REM --- Logging setup -----------------------------------------------------------
REM Build a locale-independent timestamp YYYYMMDD-HHMMSS. PowerShell is the
REM most reliable source on modern Windows (wmic was removed in Win11 24H2+);
REM fall back to wmic, then to the %DATE%/%TIME% variables.
set "TS="
where powershell.exe >nul 2>&1
if not errorlevel 1 (
    for /f "usebackq delims=" %%T in (`powershell.exe -NoProfile -Command "Get-Date -Format 'yyyyMMdd-HHmmss'"`) do set "TS=%%T"
)
if not defined TS (
    for /f "tokens=2 delims==." %%a in (
        'wmic os get LocalDateTime /value 2^>nul ^| find "="'
    ) do set "_LDT=%%a"
    if defined _LDT set "TS=!_LDT:~0,8!-!_LDT:~8,6!"
)
if not defined TS (
    set "_d=%DATE: =0%"
    set "_t=%TIME: =0%"
    set "TS=!_d:~-4!!_d:~-7,2!!_d:~-10,2!-!_t:~0,2!!_t:~3,2!!_t:~6,2!"
    set "TS=!TS::=!"
)

if not defined ARG_LOGPATH (
    set "LOG_FILE=%SCRIPT_DIR%Clear-PrintSpoolerQueue_!TS!.log"
) else (
    set "LOG_FILE=!ARG_LOGPATH!"
)

REM Touch the log file so later appends can't silently fail on a missing dir.
> "!LOG_FILE!" echo [%DATE% %TIME%] [INFO] Script v%SCRIPT_VERSION% started.
if errorlevel 1 (
    echo WARNING: Could not create log file at: !LOG_FILE!
    set "LOG_FILE="
)

REM --- Counters ----------------------------------------------------------------
set "SPL_BEFORE=0"
set "SHD_BEFORE=0"
set "TOTAL_BEFORE=0"
set "BYTES_BEFORE=0"
set "SPL_AFTER=0"
set "SHD_AFTER=0"
set "FILES_DELETED=0"
set "ERROR_COUNT=0"
set "EXIT_CODE=0"

REM --- Pretty header -----------------------------------------------------------
title Print Spooler Queue Cleanup v%SCRIPT_VERSION%
if "%FLAG_QUIET%"=="0" (
    echo ==============================================================================
    echo     Print Spooler Queue Cleanup v%SCRIPT_VERSION%
    echo     Author : Mikhail Deynekin ^<Mikhail@Deynekin.com^>
    echo     Site   : https://Deynekin.com
    echo     GitHub : https://github.com/paulmann/Print-Spooler-Queue-Cleanup
    echo ==============================================================================
    if "%FLAG_WHATIF%"=="1" echo MODE: WhatIf / dry-run -- no changes will be made.
    if "%FLAG_USEPS%"=="1"  echo MODE: Delegating to PowerShell engine.
    echo.
)
call :Log INFO "Args: Force=%FLAG_FORCE% WhatIf=%FLAG_WHATIF% UsePS=%FLAG_USEPS% Quiet=%FLAG_QUIET% Log=!LOG_FILE!"

REM --- Optional delegation to the PowerShell script ---------------------------
if "%FLAG_USEPS%"=="1" (
    if not exist "%PS_SCRIPT%" (
        call :Status ERROR "PowerShell script not found next to this batch: %PS_SCRIPT%"
        set "EXIT_CODE=1"
        goto :finish
    )
    set "PS_ARGS=-NoProfile -ExecutionPolicy Bypass -File ""%PS_SCRIPT%"""
    if "%FLAG_FORCE%"=="1"  set "PS_ARGS=!PS_ARGS! -Force"
    if "%FLAG_WHATIF%"=="1" set "PS_ARGS=!PS_ARGS! -WhatIf"
    if defined ARG_LOGPATH  set "PS_ARGS=!PS_ARGS! -LogPath ""!ARG_LOGPATH!"""
    call :Status INFO "Invoking: powershell.exe !PS_ARGS!"
    powershell.exe !PS_ARGS!
    set "EXIT_CODE=!ERRORLEVEL!"
    goto :finish
)

REM --- Admin check -------------------------------------------------------------
call :CheckAdmin
if errorlevel 1 ( set "EXIT_CODE=1" & goto :finish )

REM --- Resolve spool directory (registry override -> default) -----------------
call :ResolveSpoolDir
if errorlevel 1 ( set "EXIT_CODE=1" & goto :finish )

REM --- Pre-cleanup scan --------------------------------------------------------
call :ScanSpoolDir

REM --- Confirmation ------------------------------------------------------------
if "%FLAG_WHATIF%"=="1" (
    call :Status INFO "WhatIf: would stop %SERVICE_NAME%, delete %TOTAL_BEFORE% files (%BYTES_BEFORE% bytes), restart %SERVICE_NAME%."
    set "EXIT_CODE=0"
    goto :finish
)

if "%FLAG_FORCE%"=="0" if "%FLAG_QUIET%"=="0" (
    echo.
    set /p "_confirm=Proceed with cleanup of !TOTAL_BEFORE! file(s) in !SPOOL_DIR! ? [Y/N] "
    if /I not "!_confirm!"=="Y" (
        call :Status WARN "User declined the operation."
        set "EXIT_CODE=4"
        goto :finish
    )
)

REM --- Stop service ------------------------------------------------------------
call :StopSpooler
if errorlevel 1 ( set "EXIT_CODE=1" & goto :finish )

REM --- Delete spool files ------------------------------------------------------
call :CleanSpoolDir

REM --- Always attempt to start service, even if cleanup had errors ------------
call :StartSpooler
if errorlevel 1 set "EXIT_CODE=1"

REM --- Post-cleanup scan -------------------------------------------------------
call :PostScan

REM --- Decide final exit code --------------------------------------------------
if "%EXIT_CODE%"=="0" (
    if !ERROR_COUNT! GTR 0 (
        set "EXIT_CODE=2"
    )
)

REM --- Report ------------------------------------------------------------------
call :Report

:finish
if defined LOG_FILE call :Log INFO "Script finished with exit code !EXIT_CODE!."
if "%FLAG_QUIET%"=="0" if "%FLAG_NOPAUSE%"=="0" (
    echo.
    echo Press any key to exit . . .
    pause >nul
)
endlocal & exit /b %EXIT_CODE%


REM ============================================================================
REM  Subroutines
REM ============================================================================

:show_help
    echo.
    echo Print Spooler Queue Cleanup v%SCRIPT_VERSION%
    echo Author: Mikhail Deynekin ^<Mikhail@Deynekin.com^> -- https://Deynekin.com
    echo.
    echo USAGE:
    echo     %SCRIPT_NAME% [/Force] [/WhatIf] [/UsePS] [/Quiet] [/NoPause] [/Log ^<file^>]
    echo     %SCRIPT_NAME% /?
    echo     %SCRIPT_NAME% /Version
    echo.
    echo OPTIONS:
    echo     /Force      Skip the interactive Y/N confirmation prompt.
    echo     /WhatIf     Dry-run -- report what would happen, change nothing.
    echo     /UsePS      Delegate the cleanup to Clear-PrintSpoolerQueue.ps1
    echo                 (recommended for advanced features such as remote
    echo                 cleanup, verbose logging, and finer error reporting).
    echo     /Quiet      Suppress decorative output (logging is still written).
    echo     /NoPause    Do not pause for "press any key" before exiting.
    echo     /Log ^<file^>  Write the log to ^<file^> instead of an auto-named log.
    echo     /?          Show this help and exit.
    echo     /Version    Show script version and exit.
    echo.
    echo EXIT CODES:
    echo     0  Success
    echo     1  Fatal error (admin missing, service failure, ...)
    echo     2  Completed with warnings (Spooler restarted, but some files
    echo        could not be deleted)
    echo     3  Bad command line arguments
    echo     4  Cancelled by user
    echo.
    endlocal & exit /b 0

:show_version
    echo %SCRIPT_NAME% v%SCRIPT_VERSION%
    endlocal & exit /b 0

REM ----------------------------------------------------------------------------
REM  :Log <LEVEL> <MESSAGE>
REM    Append a single timestamped line to the log file. Never aborts the run.
REM ----------------------------------------------------------------------------
:Log
    if not defined LOG_FILE goto :eof
    set "_lvl=%~1"
    set "_msg=%~2"
    >> "!LOG_FILE!" echo [%DATE% %TIME%] [!_lvl!] !_msg!
    goto :eof

REM ----------------------------------------------------------------------------
REM  :Status <LEVEL> <MESSAGE>
REM    Print a status line to the console (unless /Quiet) AND log it.
REM    LEVEL is one of: INFO | OK | WARN | ERROR
REM ----------------------------------------------------------------------------
:Status
    set "_lvl=%~1"
    set "_msg=%~2"
    if "%FLAG_QUIET%"=="0" (
        echo [!_lvl!] !_msg!
    )
    call :Log "!_lvl!" "!_msg!"
    goto :eof

REM ----------------------------------------------------------------------------
REM  :CheckAdmin
REM    Verify elevated privileges via "net session". Exit 1 on failure.
REM ----------------------------------------------------------------------------
:CheckAdmin
    net session >nul 2>&1
    if errorlevel 1 (
        call :Status ERROR "Administrator privileges are required."
        if "%FLAG_QUIET%"=="0" (
            echo Right-click the script and choose "Run as administrator", or
            echo invoke it from an elevated cmd.exe / PowerShell session.
        )
        exit /b 1
    )
    call :Status OK "Administrator privileges confirmed."
    exit /b 0

REM ----------------------------------------------------------------------------
REM  :ResolveSpoolDir
REM    Resolve the active spool directory. Honors a custom path stored in
REM    HKLM\SYSTEM\CurrentControlSet\Control\Print\Printers!DefaultSpoolDirectory.
REM    Falls back to %SystemRoot%\System32\spool\PRINTERS.
REM ----------------------------------------------------------------------------
:ResolveSpoolDir
    set "SPOOL_DIR="
    for /f "tokens=2,*" %%A in (
        'reg query "HKLM\SYSTEM\CurrentControlSet\Control\Print\Printers" /v DefaultSpoolDirectory 2^>nul ^| find /I "DefaultSpoolDirectory"'
    ) do (
        set "SPOOL_DIR=%%B"
    )
    if not defined SPOOL_DIR set "SPOOL_DIR=%SystemRoot%\System32\spool\PRINTERS"
    REM Strip surrounding quotes if any.
    set "SPOOL_DIR=!SPOOL_DIR:"=!"

    if not exist "!SPOOL_DIR!\" (
        call :Status ERROR "Spool directory not found: !SPOOL_DIR!"
        exit /b 1
    )
    call :Status INFO "Spool directory: !SPOOL_DIR!"
    exit /b 0

REM ----------------------------------------------------------------------------
REM  :ScanSpoolDir
REM    Count *.SPL/*.SHD files and total size. Locale-independent: uses
REM    PowerShell when available (correctly sums file lengths) and falls
REM    back to a pure-cmd file count when PowerShell is missing.
REM ----------------------------------------------------------------------------
:ScanSpoolDir
    set "SPL_BEFORE=0"
    set "SHD_BEFORE=0"
    set "BYTES_BEFORE=0"

    for /f %%i in ('dir /b /a-d "!SPOOL_DIR!\*.SPL" 2^>nul ^| find /c /v ""') do set "SPL_BEFORE=%%i"
    for /f %%i in ('dir /b /a-d "!SPOOL_DIR!\*.SHD" 2^>nul ^| find /c /v ""') do set "SHD_BEFORE=%%i"
    set /a "TOTAL_BEFORE=SPL_BEFORE + SHD_BEFORE"

    REM Try PowerShell for an exact byte total (locale-independent).
    where powershell.exe >nul 2>&1
    if not errorlevel 1 (
        for /f "usebackq delims=" %%S in (`powershell.exe -NoProfile -Command "$d=Get-ChildItem -LiteralPath '!SPOOL_DIR!' -File -Force -ErrorAction SilentlyContinue ^| Where-Object { $_.Extension -ieq '.SPL' -or $_.Extension -ieq '.SHD' }; if ($d) { ($d ^| Measure-Object Length -Sum).Sum } else { 0 }"`) do (
            set "BYTES_BEFORE=%%S"
        )
    )

    if "%FLAG_QUIET%"=="0" (
        echo Pre-cleanup scan:
        echo     SPL files : !SPL_BEFORE!
        echo     SHD files : !SHD_BEFORE!
        echo     Total     : !TOTAL_BEFORE! file(s), !BYTES_BEFORE! byte(s)
        echo.
    )
    call :Log INFO "Pre-cleanup: SPL=!SPL_BEFORE! SHD=!SHD_BEFORE! Total=!TOTAL_BEFORE! Bytes=!BYTES_BEFORE!"
    exit /b 0

REM ----------------------------------------------------------------------------
REM  :StopSpooler
REM    Stop the Spooler service. Treats "service is already stopped" as success.
REM    Uses sc.exe (richer than `net stop` for state queries) but falls back
REM    to `net stop` for the actual stop call.
REM ----------------------------------------------------------------------------
:StopSpooler
    call :Status INFO "Stopping %SERVICE_NAME% service..."
    sc query %SERVICE_NAME% | find /I "STATE" | find /I "STOPPED" >nul 2>&1
    if not errorlevel 1 (
        call :Status OK "%SERVICE_NAME% was already stopped."
        exit /b 0
    )

    net stop %SERVICE_NAME% /y >nul 2>&1
    set "_rc=!ERRORLEVEL!"
    if !_rc! NEQ 0 (
        REM Re-check: a transient timing issue can return non-zero even on success.
        sc query %SERVICE_NAME% | find /I "STATE" | find /I "STOPPED" >nul 2>&1
        if not errorlevel 1 (
            call :Status OK "%SERVICE_NAME% stopped."
            exit /b 0
        )
        call :Status ERROR "Failed to stop %SERVICE_NAME% (exit code !_rc!)."
        exit /b 1
    )
    call :Status OK "%SERVICE_NAME% stopped."
    exit /b 0

REM ----------------------------------------------------------------------------
REM  :CleanSpoolDir
REM    Delete *.SHD and *.SPL files only. Counts deletions and tracks errors.
REM    Per-file deletion (rather than `del *.SPL`) allows accurate failure
REM    accounting and avoids aborting on the first locked file.
REM ----------------------------------------------------------------------------
:CleanSpoolDir
    if !TOTAL_BEFORE! EQU 0 (
        call :Status INFO "Spool directory already empty -- nothing to delete."
        exit /b 0
    )
    call :Status INFO "Deleting *.SHD and *.SPL files in !SPOOL_DIR! ..."

    for %%E in (SPL SHD) do (
        for %%F in ("!SPOOL_DIR!\*.%%E") do (
            if exist "%%~fF" (
                del /f /q "%%~fF" >nul 2>&1
                if exist "%%~fF" (
                    set /a "ERROR_COUNT+=1"
                    call :Log WARN "Could not delete: %%~fF"
                ) else (
                    set /a "FILES_DELETED+=1"
                )
            )
        )
    )

    if !ERROR_COUNT! GTR 0 (
        call :Status WARN "Deleted !FILES_DELETED! file(s); !ERROR_COUNT! could not be removed (likely locked)."
    ) else (
        call :Status OK "Deleted !FILES_DELETED! file(s)."
    )
    exit /b 0

REM ----------------------------------------------------------------------------
REM  :StartSpooler
REM    Start the Spooler service and verify it reaches RUNNING.
REM ----------------------------------------------------------------------------
:StartSpooler
    call :Status INFO "Starting %SERVICE_NAME% service..."
    net start %SERVICE_NAME% >nul 2>&1
    set "_rc=!ERRORLEVEL!"

    REM Always verify state, even when net start returns 0.
    sc query %SERVICE_NAME% | find /I "STATE" | find /I "RUNNING" >nul 2>&1
    if errorlevel 1 (
        call :Status ERROR "Failed to start %SERVICE_NAME% (net rc=!_rc!). Printing will be unavailable until the service is running."
        exit /b 1
    )
    call :Status OK "%SERVICE_NAME% is running."
    exit /b 0

REM ----------------------------------------------------------------------------
REM  :PostScan
REM    Re-count *.SPL/*.SHD after cleanup to confirm the directory is empty.
REM ----------------------------------------------------------------------------
:PostScan
    set "SPL_AFTER=0"
    set "SHD_AFTER=0"
    for /f %%i in ('dir /b /a-d "!SPOOL_DIR!\*.SPL" 2^>nul ^| find /c /v ""') do set "SPL_AFTER=%%i"
    for /f %%i in ('dir /b /a-d "!SPOOL_DIR!\*.SHD" 2^>nul ^| find /c /v ""') do set "SHD_AFTER=%%i"
    call :Log INFO "Post-cleanup: SPL=!SPL_AFTER! SHD=!SHD_AFTER!"
    exit /b 0

REM ----------------------------------------------------------------------------
REM  :Report
REM    Print a human-readable summary.
REM ----------------------------------------------------------------------------
:Report
    if "%FLAG_QUIET%"=="1" goto :eof
    echo.
    echo ==============================================================================
    echo                          OPERATION SUMMARY
    echo ==============================================================================
    echo     Spool directory  : !SPOOL_DIR!
    echo     Files before     : !TOTAL_BEFORE! (SPL=!SPL_BEFORE!, SHD=!SHD_BEFORE!)
    echo     Files after      : %SPL_AFTER% SPL + %SHD_AFTER% SHD remaining
    echo     Files deleted    : !FILES_DELETED!
    echo     Bytes processed  : !BYTES_BEFORE!
    echo     Errors           : !ERROR_COUNT!
    echo     Exit code        : !EXIT_CODE!
    if defined LOG_FILE echo     Log file         : !LOG_FILE!
    echo ==============================================================================
    goto :eof

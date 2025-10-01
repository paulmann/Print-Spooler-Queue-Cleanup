@echo off
SetLocal EnableExtensions DisableDelayedExpansion

:: ==============================================================================
:: Enhanced Print Spooler Queue Cleanup v3.0
:: ==============================================================================
:: Author: Mikhail Deynekin
:: E-Mail: mid1977@gmail.com
:: Web: https://Deynekin.COM
:: GitHub: https://github.com/mdeynekin/Print-Spooler-Queue-Cleanup
:: ------------------------------------------------------------------------------
:: Description: Safely clears stuck print jobs and resets the Print Spooler service
::              with comprehensive error handling, logging, and cross-version compatibility.
:: ==============================================================================

:: Initialize variables
set "SCRIPT_VERSION=3.0"
set "LOG_FILE=%~dpn0.log"
set "SPOOL_DIR="
set "FILES_DELETED=0"
set "SPL_DELETED=0"
set "SHD_DELETED=0"
set "SPACE_FREED=0"
set "ERROR_COUNT=0"
set "OS_MAJOR_VER=0"

:: Detect Windows version for compatibility
for /f "tokens=4-5 delims=. " %%i in ('ver') do set "OS_MAJOR_VER=%%i"
if %OS_MAJOR_VER% LSS 6 (
    set "LEGACY_MODE=1"
) else (
    set "LEGACY_MODE=0"
)

:: Main execution flow
call :Initialize
if errorlevel 1 exit /b 1

call :CheckAdminRights
if errorlevel 1 exit /b 1

call :GetSpoolerPaths
if errorlevel 1 exit /b 1

call :ScanSpoolDirectory
if errorlevel 1 exit /b 1

call :StopPrintSpooler
if errorlevel 1 exit /b 1

call :CleanSpoolDirectory
if errorlevel 1 exit /b 1

call :StartPrintSpooler
if errorlevel 1 exit /b 1

call :GenerateReport
call :Finalize
exit /b 0

:: ==============================================================================
:: FUNCTIONS
:: ==============================================================================

:Initialize
    :: Set console title
    title Print Spooler Queue Cleanup v%SCRIPT_VERSION%
    
    :: Display header
    echo ==============================================================================
    echo       Print Spooler Queue Cleanup v%SCRIPT_VERSION%
    echo ==============================================================================
    echo Author: Mikhail Deynekin
    echo GitHub: https://github.com/mdeynekin/Print-Spooler-Queue-Cleanup
    echo ------------------------------------------------------------------------------
    echo Description: Clears stuck print jobs and resets Print Spooler.
    echo ==============================================================================
    
    :: Log initialization
    call :LogMessage "INFO" "Script started (v%SCRIPT_VERSION%)"
    exit /b 0

:CheckAdminRights
    :: Verify administrative privileges
    net session >nul 2>&1
    if %errorlevel% neq 0 (
        echo.
        echo ERROR: Administrative privileges required.
        echo Please run this script as Administrator.
        echo.
        call :LogMessage "ERROR" "Administrative privileges not detected"
        exit /b 1
    )
    
    echo Administrative privileges confirmed.
    echo.
    call :LogMessage "INFO" "Administrative privileges confirmed"
    exit /b 0

:GetSpoolerPaths
    :: Get system paths dynamically
    set "SPOOL_DIR=%SystemRoot%\System32\spool\PRINTERS"
    
    :: Verify spool directory exists
    if not exist "%SPOOL_DIR%" (
        echo ERROR: Spool directory not found: %SPOOL_DIR%
        call :LogMessage "ERROR" "Spool directory not found: %SPOOL_DIR%"
        exit /b 1
    )
    
    call :LogMessage "INFO" "Spool directory: %SPOOL_DIR%"
    exit /b 0

:ScanSpoolDirectory
    :: Count files before cleanup
    set "SPL_COUNT=0"
    set "SHD_COUNT=0"
    set "TOTAL_SIZE=0"
    
    :: Count SPL files
    for /f %%i in ('dir /b "%SPOOL_DIR%\*.SPL" 2^>nul ^| find /c /v ""') do set "SPL_COUNT=%%i"
    
    :: Count SHD files
    for /f %%i in ('dir /b "%SPOOL_DIR%\*.SHD" 2^>nul ^| find /c /v ""') do set "SHD_COUNT=%%i"
    
    :: Calculate total size
    if %SPL_COUNT% gtr 0 (
        for /f "usebackq tokens=3" %%a in (`dir "%SPOOL_DIR%\*.SPL" ^| findstr /i /c:"File(s)"`) do (
            set "TOTAL_SIZE=%%a"
        )
    )
    if %SHD_COUNT% gtr 0 (
        for /f "usebackq tokens=3" %%a in (`dir "%SPOOL_DIR%\*.SHD" ^| findstr /i /c:"File(s)"`) do (
            set /a "TOTAL_SIZE+=%TOTAL_SIZE%"
        )
    )
    
    :: Display pre-cleanup status
    echo Files detected before cleanup:
    echo   SPL files: %SPL_COUNT%
    echo   SHD files: %SHD_COUNT%
    echo   Total files: %SPL_COUNT% + %SHD_COUNT% = %SPL_COUNT%
    set /a "TOTAL_FILES=%SPL_COUNT% + %SHD_COUNT%"
    echo   Total files: %TOTAL_FILES%
    
    :: Format size display
    if %TOTAL_SIZE% gtr 0 (
        set /a "SIZE_MB=%TOTAL_SIZE% / 1024"
        echo   Total size: %TOTAL_SIZE% bytes (~%SIZE_MB% MB)
    ) else (
        echo   Total size: 0 bytes
    )
    echo.
    
    call :LogMessage "INFO" "Pre-cleanup scan: %TOTAL_FILES% files (%TOTAL_SIZE% bytes)"
    exit /b 0

:StopPrintSpooler
    echo Stopping Print Spooler service...
    net stop Spooler /y >nul 2>&1
    if %errorlevel% neq 0 (
        echo ERROR: Failed to stop Print Spooler service.
        call :LogMessage "ERROR" "Failed to stop Print Spooler service (error %errorlevel%)"
        exit /b 1
    )
    
    echo Spooler service stopped successfully.
    echo.
    call :LogMessage "INFO" "Print Spooler service stopped"
    exit /b 0

:CleanSpoolDirectory
    :: Delete SPL files
    if exist "%SPOOL_DIR%\*.SPL" (
        del /q "%SPOOL_DIR%\*.SPL" >nul 2>&1
        if %errorlevel% equ 0 (
            set "SPL_DELETED=%SPL_COUNT%"
            set /a "FILES_DELETED+=%SPL_COUNT%"
        ) else (
            set /a "ERROR_COUNT+=1"
            call :LogMessage "WARNING" "Failed to delete some SPL files"
        )
    )
    
    :: Delete SHD files
    if exist "%SPOOL_DIR%\*.SHD" (
        del /q "%SPOOL_DIR%\*.SHD" >nul 2>&1
        if %errorlevel% equ 0 (
            set "SHD_DELETED=%SHD_COUNT%"
            set /a "FILES_DELETED+=%SHD_COUNT%"
        ) else (
            set /a "ERROR_COUNT+=1"
            call :LogMessage "WARNING" "Failed to delete some SHD files"
        )
    )
    
    :: Calculate space freed
    set "SPACE_FREED=%TOTAL_SIZE%"
    
    :: Display deletion results
    echo Cleaning spool directory: %SPOOL_DIR%
    echo Deleted %SPL_DELETED% SPL files.
    echo Deleted %SHD_DELETED% SHD files.
    echo.
    
    call :LogMessage "INFO" "Cleanup completed: %FILES_DELETED% files deleted (%SPACE_FREED% bytes freed)"
    exit /b 0

:StartPrintSpooler
    echo Restarting Print Spooler service...
    net start Spooler >nul 2>&1
    if %errorlevel% neq 0 (
        echo ERROR: Failed to restart Print Spooler service.
        call :LogMessage "ERROR" "Failed to restart Print Spooler service (error %errorlevel%)"
        exit /b 1
    )
    
    echo Spooler service restarted successfully.
    echo.
    call :LogMessage "INFO" "Print Spooler service restarted"
    exit /b 0

:GenerateReport
    echo ==============================================================================
    echo                      OPERATION SUMMARY
    echo ==============================================================================
    
    :: Calculate space freed in MB
    if %SPACE_FREED% gtr 0 (
        set /a "SPACE_MB=%SPACE_FREED% / 1024"
    ) else (
        set "SPACE_MB=0"
    )
    
    echo Files processed: %TOTAL_FILES%
    echo Files deleted:   %FILES_DELETED%
    echo SPL files:       %SPL_DELETED%
    echo SHD files:       %SHD_DELETED%
    echo Space freed:     %SPACE_FREED% bytes (~%SPACE_MB% MB)
    echo Errors:          %ERROR_COUNT%
    echo.
    
    call :LogMessage "INFO" "Operation summary - Files: %FILES_DELETED%/%TOTAL_FILES%, Space: %SPACE_FREED% bytes, Errors: %ERROR_COUNT%"
    exit /b 0

:Finalize
    call :LogMessage "INFO" "Script completed successfully"
    if %LEGACY_MODE% equ 0 (
        timeout /t 5 >nul
    ) else (
        ping -n 6 127.0.0.1 >nul
    )
    exit /b 0

:LogMessage
    set "LEVEL=%~1"
    set "MESSAGE=%~2"
    echo [%date% %time:~0,8%] [%LEVEL%] %MESSAGE% >> "%LOG_FILE%"
    exit /b 0

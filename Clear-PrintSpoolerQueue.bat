@echo off
:: ----------------------------------------------------------------------------
:: File: Clear-PrintSpoolerQueue.bat
:: Description:
::   Forcibly clears the Windows Print Spooler queue on all printers.
::   - Verifies administrative privileges
::   - Stops the Print Spooler service
::   - Deletes all spool files (*.SPL, *.SHD)
::   - Restarts the Print Spooler service
:: Author: Mikhail Deynekin
:: GitHub: https://github.com/mdeynekin/print-spooler-cleanup
:: ----------------------------------------------------------------------------

:: Check for administrative privileges
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo ERROR: Administrator privileges required.
    echo Please run this script as Administrator.
    pause
    exit /b 1
)

:: Define variables
set "SERVICE_NAME=Spooler"
set "SPOOL_DIR=%SystemRoot%\System32\spool\PRINTERS"

:: Stop the Print Spooler service
echo Stopping %SERVICE_NAME% service...
net stop "%SERVICE_NAME%" >nul 2>&1
if errorlevel 1 (
    echo ERROR: Failed to stop %SERVICE_NAME% service.
    exit /b 1
)

:: Wait to ensure service has stopped
timeout /t 3 /nobreak >nul

:: Delete spool files
echo Clearing spooler directory: %SPOOL_DIR%
del /F /Q "%SPOOL_DIR%\*.spl"  >nul 2>&1
del /F /Q "%SPOOL_DIR%\*.shd"  >nul 2>&1

:: Restart the Print Spooler service
echo Restarting %SERVICE_NAME% service...
net start "%SERVICE_NAME%" >nul 2>&1
if errorlevel 1 (
    echo ERROR: Failed to restart %SERVICE_NAME% service.
    exit /b 1
)

echo Print Spooler queue cleared and service restarted successfully.
exit /b 0

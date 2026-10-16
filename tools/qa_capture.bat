@echo off
rem ---------------------------------------------------------------------------
rem  QA self-capture: renders EVERY screen (wizard pages, main, settings,
rem  restore, about) in BOTH themes to ..\qa-shots\ using the built-in
rem  PSQC_SHOTDIR hook. Run on Windows 10/11. No interaction needed.
rem ---------------------------------------------------------------------------
setlocal
cd /d "%~dp0..\dist"
set SHOTS=%~dp0..\qa-shots
if not exist "%SHOTS%" mkdir "%SHOTS%"
del /q "%SHOTS%\*.bmp" 2>nul
set PSQC_SHOTDIR=%SHOTS%
echo [1/2] light theme capture...
del /q "%APPDATA%\SpoolerCleanup\settings.json" 2>nul
SpoolerCleanup.exe /theme=light
echo [2/2] dark theme capture...
del /q "%APPDATA%\SpoolerCleanup\settings.json" 2>nul
SpoolerCleanup.exe /theme=dark
set PSQC_SHOTDIR=
echo Done. Screenshots: %SHOTS%
explorer "%SHOTS%"

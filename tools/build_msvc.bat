@echo off
rem ---------------------------------------------------------------------------
rem  Print Spooler Queue Cleanup v5 - MSVC build (x64, static, production)
rem  Run from an "x64 Native Tools Command Prompt for VS 2022+".
rem ---------------------------------------------------------------------------
setlocal enabledelayedexpansion
cd /d "%~dp0.."
if not exist dist mkdir dist

set CXXFLAGS=/nologo /std:c++latest /W4 /utf-8 /O2 /MT /EHsc /GS /guard:cf /DYNAMICBASE /NXCOMPAT /HIGHENTROPYVA /DNDEBUG /DUNICODE /D_UNICODE /DWIN32_LEAN_AND_MEAN /DNOMINMAX /D_WIN32_WINNT=0x0A00 /Isrc
set LDFLAGS=/nologo /SUBSYSTEM:WINDOWS /LTCG /OPT:REF,ICF /manifest:embed
rem Store build: add /DPSQC_STORE_BUILD to CXXFLAGS (disables in-app updater)

echo [rc]  res\app.rc
rc /nologo /fo dist\app.res res\app.rc || goto :err

echo [cxx] src\*.cpp
set SRC=
for %%f in (src\*.cpp) do set SRC=!SRC! %%f
cl %CXXFLAGS% /c !SRC! /Fodist\ || goto :err

set OBJ=
for %%f in (dist\*.obj) do set OBJ=!OBJ! %%f
echo [link] dist\SpoolerCleanup.exe
link %LDFLAGS% !OBJ! dist\app.res /OUT:dist\SpoolerCleanup.exe ^
     comctl32.lib shell32.lib shlwapi.lib uxtheme.lib dwmapi.lib winspool.lib winhttp.lib ^
     advapi32.lib ole32.lib uuid.lib oleaut32.lib gdi32.lib user32.lib kernel32.lib || goto :err

del /q dist\*.obj dist\app.res 2>nul
echo [ok]  dist\SpoolerCleanup.exe
exit /b 0
:err
echo [FAIL] build failed
exit /b 1

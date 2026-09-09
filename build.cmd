@echo off
rem DeskTiler command-line build script
rem Requires Embarcadero Delphi 10.3 (Studio 19.0). Adjust BDS if installed elsewhere.

setlocal
set "BDS=g:\Program Files (x86)\Embarcadero\Studio\19.0"
set "DCC=%BDS%\bin\dcc32.exe"
set "BRCC=%BDS%\bin\brcc32.exe"
set "RELEASE=%BDS%\lib\win32\release"
cd /d "%~dp0"

echo [1/4] Generating icon DeskTiler.ico ...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\makeicon.ps1"
if errorlevel 1 goto :err

echo [2/4] Compiling resources DeskTiler.rc ...
"%BRCC%" -foDeskTiler.res "DeskTiler.rc"
if errorlevel 1 goto :err

echo [3/4] Compiling DeskTiler.dpr (dcc32, Win32 GUI) ...
"%DCC%" -Q -B -U"%RELEASE%" "DeskTiler.dpr"
if errorlevel 1 goto :err

echo.
echo [4/4] Done: %~dp0DeskTiler.exe
exit /b 0

:err
echo.
echo BUILD FAILED.
exit /b 1

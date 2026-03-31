@echo off
REM ASCM Installer Launcher
REM This batch file runs the PowerShell installer with proper execution policy

setlocal enabledelayedexpansion

REM Get the directory where this batch file is located
set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%Install-ASCM.ps1"

REM Check if the PowerShell script exists
if not exist "%PS_SCRIPT%" (
    color 4F
    echo.
    echo ERROR: Install-ASCM.ps1 not found!
    echo Expected location: %PS_SCRIPT%
    echo.
    pause
    exit /b 1
)

REM Run the PowerShell installer
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%"

REM Exit
exit /b %errorlevel%

@echo off
REM Scan Software Installer
REM This script installs Scan Software to Program Files and creates a shortcut

setlocal enabledelayedexpansion

REM Check if running as administrator
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo This installer requires administrator privileges.
    echo Please run this file as Administrator.
    echo.
    echo To run as Administrator:
    echo 1. Right-click this file
    echo 2. Select "Run as administrator"
    pause
    exit /b 1
)

REM Define paths
set SOURCE_EXE="%~dp0dist\Scan Software.exe"
set INSTALL_DIR=%ProgramFiles%\Scan Software
set TARGET_EXE=!INSTALL_DIR!\Scan Software.exe
set SHORTCUT=%ProgramData%\Microsoft\Windows\Start Menu\Programs\Scan Software.lnk

REM Create installation directory
if not exist "!INSTALL_DIR!" (
    mkdir "!INSTALL_DIR!"
    echo Created directory: !INSTALL_DIR!
)

REM Copy exe to installation directory
if exist !SOURCE_EXE! (
    copy !SOURCE_EXE! "!TARGET_EXE!" /Y
    echo Successfully installed: !TARGET_EXE!
) else (
    echo Error: Could not find !SOURCE_EXE!
    pause
    exit /b 1
)

REM Create shortcut using VBScript
set VBS_FILE="%TEMP%\CreateShortcut.vbs"

(
    echo Set oWS = WScript.CreateObject("WScript.Shell"^)
    echo sLinkFile = "!SHORTCUT!"
    echo Set oLink = oWS.CreateShortcut(sLinkFile^)
    echo oLink.TargetPath = "!TARGET_EXE!"
    echo oLink.WorkingDirectory = "!INSTALL_DIR!"
    echo oLink.Description = "Scan Software - Compliance Scanner"
    echo oLink.IconLocation = "!TARGET_EXE!"
    echo oLink.Save
) > !VBS_FILE!

cscript !VBS_FILE!
del !VBS_FILE!

echo.
echo ============================================================
echo Installation Complete!
echo ============================================================
echo.
echo Scan Software has been installed to:
echo !INSTALL_DIR!
echo.
echo You can now launch it from:
echo - Start Menu: "Scan Software"
echo - File Explorer: !INSTALL_DIR!
echo.
echo ============================================================
pause

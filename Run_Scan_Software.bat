@echo off
REM Scan Software Launcher
REM This script copies and runs from user temp folder (no admin rights needed)

setlocal enabledelayedexpansion

REM Define paths
set SOURCE_EXE="%~dp0dist\Scan Software.exe"
set TEMP_DIR=%TEMP%
set TARGET_EXE=!TEMP_DIR!\Scan Software.exe

echo ============================================================
echo  SCAN SOFTWARE LAUNCHER
echo ============================================================
echo.

REM Check if source exe exists
if not exist !SOURCE_EXE! (
    echo Error: Could not find Scan Software.exe
    echo Expected location: !SOURCE_EXE!
    echo.
    pause
    exit /b 1
)

echo Preparing to run Scan Software...
echo.

REM Copy exe to temp folder
echo Copying software to temporary folder...
copy !SOURCE_EXE! "!TARGET_EXE!" /Y >nul 2>&1

REM Check if copy was successful
if not exist "!TARGET_EXE!" (
    echo Error: Could not copy to temp folder
    echo This might be a permissions issue
    echo.
    pause
    exit /b 1
)

echo Running software from temporary location...
echo.

REM Run the exe from temp folder
start "" "!TARGET_EXE!"

REM Inform user
echo Software launched!
echo.
echo The compliance report will open automatically after scanning completes.
echo.
pause

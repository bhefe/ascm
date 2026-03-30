@echo off
REM Scan Software - Download and Run
REM Single file solution - downloads exe and runs from temp folder
REM No admin rights needed

setlocal enabledelayedexpansion

REM GitHub raw file URL
set "GITHUB_URL=https://github.com/bhefe/ascm/raw/main/dist/Scan Software.exe"
set "TEMP_DIR=%TEMP%"
set "TARGET_EXE=%TEMP_DIR%\Scan Software.exe"

echo ============================================================
echo  SCAN SOFTWARE DOWNLOADER
echo ============================================================
echo.
echo Downloading from GitHub...
echo.

REM Remove old version if exists
if exist "!TARGET_EXE!" (
    del "!TARGET_EXE!" /Q >nul 2>&1
)

REM Download using certutil (built-in Windows tool, no dependencies)
certutil -urlcache -split -f "!GITHUB_URL!" "!TARGET_EXE!"

REM Check if download was successful
if not exist "!TARGET_EXE!" (
    echo.
    echo Error: Could not download Scan Software.exe
    echo.
    echo Possible reasons:
    echo - Internet connection issue
    echo - GitHub URL is not accessible
    echo - Proxy/firewall blocking
    echo.
    echo Try again in a moment or check your internet connection.
    echo.
    pause
    exit /b 1
)

echo.
echo Download complete!
echo.
echo Running software...
echo.

REM Run the exe from temp folder
"%TARGET_EXE%"

echo.
echo Software launched!
echo.
pause

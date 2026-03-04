@echo off
title Software Scan
cd /d "%~dp0"

if exist "Scan Software.exe" (
    "Scan Software.exe"
    goto :done
)

if exist "dist\Scan Software.exe" (
    "dist\Scan Software.exe"
    goto :done
)

echo.
echo   ERROR: Scan Software.exe not found.
echo   Please run: python build_scan_exe.py
echo.
pause
exit /b 1

:done

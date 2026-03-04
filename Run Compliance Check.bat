@echo off
title Software Compliance Check
cd /d "%~dp0"

if exist "Software Compliance Check.exe" (
    "Software Compliance Check.exe"
    goto :done
)

python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo   "Software Compliance Check.exe" not found.
    echo.
    pause
    exit /b 1
)

pip install pandas pdfplumber openpyxl >nul 2>&1
python run_compliance.py

:done
pause

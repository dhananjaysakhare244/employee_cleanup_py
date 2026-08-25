@echo off
setlocal
cd /d "%~dp0"

echo ==========================================
echo Employee Cleanup Tool - Setup
echo ==========================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found.
    echo Install Python 3 from https://www.python.org/downloads/windows/
    echo Make sure "Add python.exe to PATH" is selected.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment...
    python -m venv .venv
)

echo Installing dependencies...
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements.txt

echo.
echo Starting Employee Cleanup Tool...
.venv\Scripts\python.exe main.py

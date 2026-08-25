@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Virtual environment not found.
    echo Run setup_and_run.bat first.
    pause
    exit /b 1
)

echo Installing/updating dependencies...
.venv\Scripts\python.exe -m pip install -r requirements.txt

echo Building EXE...
.venv\Scripts\python.exe -m PyInstaller --clean --onefile --windowed --name EmployeeCleanupTool main.py

echo.
echo EXE created at:
echo %CD%\dist\EmployeeCleanupTool.exe
pause

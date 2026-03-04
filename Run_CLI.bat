@echo off
REM Quick launcher for command-line version
REM This runs the script automatically to detect which step to perform

title D365 vs SafeContractor Status Comparison Tool - Auto Mode

echo.
echo ============================================================
echo  D365 vs SafeContractor Status Comparison
echo ============================================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    pause
    exit /b 1
)

REM Check dependencies
python -c "import pandas, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Installing dependencies...
    python -m pip install -r requirements.txt
)

REM Run the script
python main.py

echo.
echo ============================================================
echo  Process Complete
echo ============================================================
echo.
pause

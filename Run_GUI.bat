@echo off
REM D365 vs SafeContractor Status Comparison - Launcher
REM This batch file launches the GUI application

title D365 Status Comparison Tool

echo.
echo ============================================================
echo  D365 vs SafeContractor Status Comparison Tool
echo ============================================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo.
    echo Please install Python from: https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
)

echo [OK] Python is installed
echo.

REM Check if required packages are installed
echo Checking dependencies...
python -c "import pandas, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo.
    echo [INFO] Installing required packages...
    echo.
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo ERROR: Failed to install dependencies
        echo.
        pause
        exit /b 1
    )
    echo.
    echo [OK] Dependencies installed successfully
    echo.
) else (
    echo [OK] All dependencies are installed
    echo.
)

REM Launch the GUI application
echo Launching GUI application...
echo.
python gui_app.py

REM If the GUI exits with an error, pause to show the error
if errorlevel 1 (
    echo.
    echo ============================================================
    echo  Application closed with an error
    echo ============================================================
    echo.
    pause
)

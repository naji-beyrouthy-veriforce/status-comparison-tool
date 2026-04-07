@echo off
REM D365 vs SafeContractor Status Comparison - Launcher
REM This batch file launches the GUI application

title Status Comparison Tool

REM Set Redash API key for automated mode
set REDASH_API_KEY=RpWSRcBbV8IHkvXumk442ttCiU2j9XLSa0niHXRD

REM -----------------------------------------------------------------------
REM  D365 API credentials — fill in once IT provides the Azure App Registration
REM  When all three are set, D365 files are downloaded automatically.
REM -----------------------------------------------------------------------
REM set D365_TENANT_ID=YOUR_TENANT_ID_HERE
REM set D365_CLIENT_ID=YOUR_CLIENT_ID_HERE
REM set D365_CLIENT_SECRET=YOUR_CLIENT_SECRET_HERE

REM  D365 saved view IDs (confirmed April 2026 — hardcoded in config.py, no need to re-set here)
REM  Uncomment only if you need to temporarily override a specific view.
REM set D365_VIEW_ID_ACCREDITATION=2102f6c1-4411-f011-998a-000d3ab02833
REM set D365_VIEW_ID_WCB=06e9e4df-4411-f011-998a-000d3ab02833
REM set D365_VIEW_ID_CLIENT=4b79190b-4511-f011-998a-000d3ab02833
REM set D365_VIEW_ID_CRITICAL_DOCUMENT=a007b506-6e27-f111-8342-7ced8d421558
REM set D365_VIEW_ID_ESG=990883d8-4b28-f111-8342-0022489c5458

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

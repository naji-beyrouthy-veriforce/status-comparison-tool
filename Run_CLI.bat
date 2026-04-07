@echo off
REM Quick launcher for command-line version
REM This runs the script automatically to detect which step to perform

title D365 vs SafeContractor Status Comparison Tool - Auto Mode

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

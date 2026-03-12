@echo off
REM Batch file to generate email report from comparison Excel files

echo ===============================================
echo    Email Report Generator
echo ===============================================
echo.

REM Run the email report generator
python -m src.email_report

echo.
echo ===============================================
echo Press any key to exit...
pause > nul

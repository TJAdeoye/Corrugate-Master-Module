@echo off
REM ===================================================================
REM  Run Dashboards.bat - rebuild every Corrugate Master Module
REM  dashboard, then open the hub page.
REM
REM  Keep this file in the Corrugate Master Module folder.
REM  Double-click it, or point Task Scheduler at it for a daily run.
REM ===================================================================

cd /d "%~dp0"

echo.
echo  Corrugate Master Module - rebuilding
echo  ====================================
echo.

python build_all.py
set RESULT=%ERRORLEVEL%

echo.
if %RESULT% NEQ 0 (
    echo  ONE OR MORE DASHBOARDS FAILED - see the messages above.
    echo  The ones that built are still up to date.
    echo.
    pause
    exit /b %RESULT%
)

echo  Opening the hub...
start "" "Corrugate Master Module.html"

REM Comment out the next line if you would rather the window close itself.
pause

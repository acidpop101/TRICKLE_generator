@echo off
setlocal
set "BASE_DIR=%~dp0"
set "PYTHONPATH=%BASE_DIR%src"
set "OUTPUT_DIR=%BASE_DIR%cbs_output"

echo ========================================================
echo                 CLEAR ATM DISPUTES DATA
echo ========================================================
echo WARNING: This action will permanently delete all logged
echo entries from the 'disputes' sheet in atm_data.xlsx.
echo It will also clear the old files in the cbs_output folder.
echo.
echo Make sure Microsoft Excel is completely CLOSED before
echo you proceed.
echo ========================================================
echo.
pause

echo.
echo Cleaning Excel data, please wait...
python "%BASE_DIR%src\atm_dispute_app\clear_disputes.py"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Script exited with error code %ERRORLEVEL%
)

echo.
echo Cleaning output directory %OUTPUT_DIR%...
if exist "%OUTPUT_DIR%" (
    del /Q /S "%OUTPUT_DIR%\*.txt" >nul 2>&1
    echo [SUCCESS] cbs_output cleared.
)

REM Cleanup Python __pycache__ just in case
for /d /r "%BASE_DIR%" %%d in (__pycache__) do @if exist "%%d" rd /s /q "%%d"

echo.
echo ========================================================
echo                  ALL CLEARED!
echo ========================================================
pause
endlocal

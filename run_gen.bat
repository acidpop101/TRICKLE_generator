@echo off
REM Run Trickle Feed Generator (Excel -> Text)
REM Uses valid relative paths

set "BASE_DIR=%~dp0"
set "TRICKLE_DIR=%BASE_DIR%trickle_feed"
set "APP_DIR=%BASE_DIR%atm_dispute_app"

echo Clean old output...
del /Q "%TRICKLE_DIR%\cbs_output\*.txt"

echo Running Generator...
python "%TRICKLE_DIR%\main.py" "%APP_DIR%\atm_data.xlsx" "%TRICKLE_DIR%\cbs_output"

echo.
echo Check the output folder: "%TRICKLE_DIR%\cbs_output"
pause

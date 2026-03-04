@echo off
setlocal
set "BASE_DIR=%~dp0"
set "PYTHONPATH=%BASE_DIR%src"

echo ========================================================
echo  Running Trickle Generator (v3.1 Portable)
echo ========================================================

REM Check for Python
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python is not installed or not in PATH.
    echo Please install Python [3.7 or newer] to run this application.
    pause
    exit /b 1
)

REM Check dependencies
python -c "import openpyxl" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    if exist "%BASE_DIR%requirements.txt" (
        echo [INFO] Installing dependencies...
        pip install -r "%BASE_DIR%requirements.txt"
    ) else (
        echo [ERROR] Missing requirements.txt. Cannot run.
        pause
        exit /b 1
    )
)

set "GENERATOR_SCRIPT=%BASE_DIR%src\trickle_feed\main.py"
set "INPUT_EXCEL=%BASE_DIR%data\atm_data.xlsx"
set "OUTPUT_DIR=%BASE_DIR%cbs_output"

echo  Base Dir: %BASE_DIR%
echo  Input:    %INPUT_EXCEL%
echo  Output:   %OUTPUT_DIR%
echo ========================================================

REM Check Input
if not exist "%INPUT_EXCEL%" (
    echo [ERROR] Input file not found at:
    echo %INPUT_EXCEL%
    echo Please ensure 'atm_data.xlsx' is inside the 'data' folder.
    pause
    exit /b 1
)

REM Check Script
if not exist "%GENERATOR_SCRIPT%" (
    echo [ERROR] Python script not found at:
    echo %GENERATOR_SCRIPT%
    pause
    exit /b 1
)

REM Clean Output
if exist "%OUTPUT_DIR%" (
    echo Cleaning old files...
    del /Q "%OUTPUT_DIR%\*.txt"
) else (
    mkdir "%OUTPUT_DIR%"
)

echo.
echo Processing...
python "%GENERATOR_SCRIPT%" "%INPUT_EXCEL%" "%OUTPUT_DIR%"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Generation failed.
) else (
    echo.
    echo [SUCCESS] Files generated in: %OUTPUT_DIR%
)

endlocal
pause

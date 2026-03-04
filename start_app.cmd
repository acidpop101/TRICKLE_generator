@echo off
setlocal
set "BASE_DIR=%~dp0"
set "PYTHONPATH=%BASE_DIR%src"

REM Check for Python
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python is not installed or not in PATH.
    echo Please install Python [3.7 or newer] to run this application.
    pause
    exit /b 1
)

REM Check dependencies (simple check for openpyxl)
python -c "import openpyxl" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [INFO] Dependencies not found. Installing from requirements.txt...
    if exist "%BASE_DIR%requirements.txt" (
        pip install -r "%BASE_DIR%requirements.txt"
    ) else (
        echo [ERROR] requirements.txt not found!
        echo Please ensure requirements.txt is in the project root.
        pause
        exit /b 1
    )
)

echo Starting ATM Dispute App...
python "%BASE_DIR%src\atm_dispute_app\ATM_DISPUTE.py"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Application exited with error code %ERRORLEVEL%
    pause
)
endlocal

@echo off
SETLOCAL EnableDelayedExpansion
TITLE NIPT Report Generator
cd /d "%~dp0"

echo.
echo ===========================================
echo   NIPT Report Generator
echo ===========================================
echo.

REM Check for venv
if not exist ".venv" (
    echo [!] Virtual environment not found. Running setup...
    call setup_nipt.bat
    if errorlevel 1 (
        echo [ERROR] Setup failed.
        pause
        exit /b 1
    )
)

echo [*] Activating environment...
call .venv\Scripts\activate.bat
if errorlevel 1 (
    echo [ERROR] Failed to activate virtual environment.
    pause
    exit /b 1
)

echo [*] Starting Application...
python nipt_report_generator.py 2> launch_crash.log
if errorlevel 1 (
    echo.
    echo [CRITICAL ERROR] Application crashed or failed to start.
    echo --------------------------------------------------
    type launch_crash.log
    echo --------------------------------------------------
    echo.
    echo Please send the above error message to support.
    pause
) else (
    echo [OK] Application closed normally.
)

REM Keep window open if it closed too fast
timeout /t 5

@echo off
TITLE NIPT Report Generator
cd /d "%~dp0"

if not exist ".venv" (
    echo [!] Virtual environment not found. Running setup...
    call setup_nipt.bat
)

echo Starting NIPT Report Generator...
call .venv\Scripts\activate.bat
python nipt_report_generator.py
if errorlevel 1 (
    echo.
    echo [ERROR] Application crashed.
    pause
)

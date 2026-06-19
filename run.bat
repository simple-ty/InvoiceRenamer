@echo off
chcp 65001 >nul
REM Invoice Renamer - Startup Script

set VENV_PYTHON=C:\Users\29292\.workbuddy\binaries\python\envs\invoice-renamer\Scripts\python.exe

if not exist "%VENV_PYTHON%" (
    echo Creating virtual environment...
    C:\Users\29292\.workbuddy\binaries\python\versions\3.13.12\python.exe -m venv C:\Users\29292\.workbuddy\binaries\python\envs\invoice-renamer
    echo Installing dependencies...
    C:\Users\29292\.workbuddy\binaries\python\envs\invoice-renamer\Scripts\pip.exe install pdfplumber customtkinter openpyxl
)

"%VENV_PYTHON%" "%~dp0invoice_renamer_ui.py" %*

echo.
echo Program exited. Press any key to close...
pause >nul

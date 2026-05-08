@echo off
echo.
echo =============================================
echo   TestRail - Set Run Dates Setup
echo =============================================
echo.

cd /d "%~dp0"

echo [1/3] Checking .env file...
if not exist .env (
    echo   ERROR: .env file not found.
    echo          Please create .env file and try again.
    pause
    exit /b 0
)
echo   OK - .env file found

echo.
echo [2/3] Checking Python...
python --version > nul 2>&1
if errorlevel 1 (
    echo   ERROR: Python not found.
    echo          Install from https://www.python.org/downloads
    pause
    exit /b 1
)
for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo   OK - %%v

echo.
echo [3/3] Installing libraries and building set_run_dates.exe...
echo   Please wait 1~3 minutes...
echo.

if exist venv\Scripts\pip.exe (
    venv\Scripts\pip.exe install requests python-dotenv pyinstaller --quiet > nul 2>&1
) else (
    python -m pip install requests python-dotenv pyinstaller --quiet > nul 2>&1
)

if exist venv\Scripts\pyinstaller.exe (
    set PYINST=venv\Scripts\pyinstaller.exe
) else (
    set PYINST=python -m PyInstaller
)

if exist set_run_dates.exe del /q set_run_dates.exe
if exist build rmdir /s /q build
if exist set_run_dates.spec del /q set_run_dates.spec

%PYINST% --onefile --windowed --distpath . --hidden-import dotenv --hidden-import requests --hidden-import urllib3 --hidden-import tkinter --hidden-import tkinter.ttk --hidden-import calendar set_run_dates.py > nul 2>&1

if not exist set_run_dates.exe (
    echo   ERROR: Build failed.
    pause
    exit /b 1
)

if exist build rmdir /s /q build
if exist set_run_dates.spec del /q set_run_dates.spec

echo   OK - set_run_dates.exe done
echo.
echo =============================================
echo   Build Complete!
echo.
echo   set_run_dates.exe  - Run this to set dates
echo   (keep .env in the same folder)
echo =============================================
echo.
pause
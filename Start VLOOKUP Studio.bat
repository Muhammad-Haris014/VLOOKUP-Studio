@echo off
title VLOOKUP Studio
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found. Install Python 3.11 or newer from https://www.python.org/downloads/
    echo and tick "Add python.exe to PATH" during the install.
    pause
    exit /b 1
)

python -m pip install -q -r requirements.txt
if errorlevel 1 (
    echo Could not install the required packages. Check the internet connection and try again.
    pause
    exit /b 1
)

rem Opens the browser itself and moves to the next free port if 8501 is busy.
python -m streamlit run app.py --server.headless false
pause

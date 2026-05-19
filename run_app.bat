@echo off
setlocal

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo.
    echo TMC Processor is not set up yet.
    echo Please double-click setup_windows.bat first.
    echo.
    pause
    exit /b 1
)

echo.
echo Starting TMC Processor ...
echo A browser window should open automatically.
echo Press Ctrl+C in this window to stop the app.
echo.

".venv\Scripts\python.exe" -m streamlit run app.py

echo.
pause

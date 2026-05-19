@echo off
setlocal

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo.
    echo TMC Processor is not set up yet.
    echo Please double-click start_tmc_processor.bat or setup_windows.bat first.
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
if errorlevel 1 (
    echo.
    echo TMC Processor could not be started.
    echo.
    pause
    exit /b 1
)

exit /b 0

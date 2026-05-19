@echo off
setlocal

cd /d "%~dp0"

echo.
echo TMC Processor - Windows setup
echo =============================
echo.

where python >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_CMD=python"
) else (
    where py >nul 2>nul
    if errorlevel 1 (
        echo Python was not found.
        echo Please install Python 3.10 or newer, then run this file again.
        echo Download Python from https://www.python.org/downloads/windows/
        echo.
        pause
        exit /b 1
    )
    set "PYTHON_CMD=py -3"
)

%PYTHON_CMD% --version
if errorlevel 1 (
    echo.
    echo Python could not be started.
    pause
    exit /b 1
)

echo.
echo Creating virtual environment in .venv ...
%PYTHON_CMD% -m venv .venv
if errorlevel 1 (
    echo.
    echo Failed to create the virtual environment.
    pause
    exit /b 1
)

echo.
echo Installing TMC Processor and required packages ...
call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
if errorlevel 1 (
    echo.
    echo Failed to upgrade pip.
    pause
    exit /b 1
)

python -m pip install -e .
if errorlevel 1 (
    echo.
    echo Failed to install TMC Processor.
    pause
    exit /b 1
)

echo.
echo Setup complete.
echo Double-click run_app.bat to start TMC Processor.
echo.
pause

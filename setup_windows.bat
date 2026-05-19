@echo off
setlocal

cd /d "%~dp0"

echo.
echo TMC Processor - Windows setup
echo =============================
echo.

if exist ".venv\" goto CHECK_VENV

echo First-time setup is required.
echo This may take several minutes. Please keep this window open.
echo.

call :FIND_PYTHON
if errorlevel 1 goto ERROR_END

%PYTHON_CMD% --version
if errorlevel 1 (
    echo.
    echo Python was found but could not be started.
    goto ERROR_END
)

echo.
echo Creating virtual environment in .venv ...
%PYTHON_CMD% -m venv .venv
if errorlevel 1 (
    echo.
    echo Failed to create .venv.
    goto ERROR_END
)

:CHECK_VENV
if not exist ".venv\Scripts\python.exe" (
    echo.
    echo The .venv folder exists, but .venv\Scripts\python.exe was not found.
    echo Delete the .venv folder and run this file again.
    goto ERROR_END
)

echo.
echo Upgrading pip ...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 (
    echo.
    echo Failed to upgrade pip.
    goto ERROR_END
)

echo.
echo Installing TMC Processor and required packages ...
".venv\Scripts\python.exe" -m pip install -e .
if errorlevel 1 (
    echo.
    echo Failed to install TMC Processor dependencies.
    goto ERROR_END
)

echo.
echo Installing pywin32 for Excel COM support ...
".venv\Scripts\python.exe" -m pip install pywin32
if errorlevel 1 (
    echo.
    echo Warning: pywin32 installation failed.
    echo The app can still run, but Excel Template Mode may be unavailable.
) else (
    echo pywin32 installed.
)

echo.
echo Optional Excel COM check ...
".venv\Scripts\python.exe" -c "import win32com.client as client; excel = client.Dispatch('Excel.Application'); print('Excel COM available. Excel version:', excel.Version); excel.Quit()"
if errorlevel 1 (
    echo.
    echo Warning: Excel COM is not available on this machine.
    echo The app can still run using Safe PNG Export Mode.
)

echo.
echo Setup complete.
echo Double-click run_app.bat to start TMC Processor.
echo.
exit /b 0

:FIND_PYTHON
python --version >nul 2>nul
if not errorlevel 1 (
    set "PYTHON_CMD=python"
    exit /b 0
)

py -3 --version >nul 2>nul
if not errorlevel 1 (
    set "PYTHON_CMD=py -3"
    exit /b 0
)

echo Python was not found.
echo Please install Python 3.10 or newer, then run this file again.
echo Download Python from: https://www.python.org/downloads/windows/
echo Important: during installation, select "Add python.exe to PATH".
exit /b 1

:ERROR_END
echo.
echo Setup stopped because of the error above.
echo.
pause
exit /b 1

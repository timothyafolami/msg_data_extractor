@echo off
setlocal enabledelayedexpansion

set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
set "VENV_DIR=%SCRIPT_DIR%\.venv"

set "PYTHON_BIN="
where py >nul 2>nul
if %ERRORLEVEL%==0 (
    set "PYTHON_BIN=py -3"
) else (
    where python >nul 2>nul
    if %ERRORLEVEL%==0 (
        set "PYTHON_BIN=python"
    )
)

if not defined PYTHON_BIN (
    echo Python 3 is required but was not found on PATH.
    exit /b 1
)

echo Using Python: %PYTHON_BIN%

if not exist "%VENV_DIR%\Scripts\python.exe" (
    echo Creating virtual environment at "%VENV_DIR%"
    call %PYTHON_BIN% -m venv "%VENV_DIR%"
    if errorlevel 1 exit /b 1
) else (
    echo Virtual environment already exists at "%VENV_DIR%"
)

echo Upgrading pip
call "%VENV_DIR%\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 exit /b 1

echo Installing dependencies
call "%VENV_DIR%\Scripts\python.exe" -m pip install -r "%SCRIPT_DIR%\requirements.txt"
if errorlevel 1 exit /b 1

echo.
echo Setup complete.
echo Run processing with:
echo   process_folder.bat P:\path\to\input_folder [P:\path\to\output_folder]
exit /b 0

@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
set "VENV_PYTHON=%SCRIPT_DIR%\.venv\Scripts\python.exe"

if "%~1"=="" goto usage
if /I "%~1"=="-h" goto usage
if /I "%~1"=="--help" goto usage

if not exist "%VENV_PYTHON%" (
    echo Virtual environment not found. Run setup_local.bat first.
    exit /b 1
)

set "INPUT_FOLDER=%~1"
set "OUTPUT_FOLDER=%~2"

if not exist "%INPUT_FOLDER%" (
    echo Input folder does not exist: "%INPUT_FOLDER%"
    exit /b 1
)

if "%OUTPUT_FOLDER%"=="" (
    for %%I in ("%INPUT_FOLDER%") do set "OUTPUT_FOLDER=%SCRIPT_DIR%\%%~nxI_extracted"
)

call "%VENV_PYTHON%" "%SCRIPT_DIR%\extract_msg_photos.py" "%INPUT_FOLDER%" --output-folder "%OUTPUT_FOLDER%"
exit /b %ERRORLEVEL%

:usage
echo Usage:
echo   process_folder.bat ^<input-folder^> [output-folder]
echo.
echo Examples:
echo   process_folder.bat P:\client_folder
echo   process_folder.bat P:\client_folder P:\output_folder
exit /b 1
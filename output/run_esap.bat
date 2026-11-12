@echo off
rem eSAP: build the draft SAP from the template and the study metadata.
rem   Double-click this file, or run:   run_esap.bat [config.toml] [--check-only]
rem The first run creates a private Python environment in the .venv folder here and
rem installs the packages listed in requirements.txt (needs internet once).
rem Later runs go straight to the script.
setlocal EnableExtensions
cd /d "%~dp0"

set "CONFIG=%~1"
if "%CONFIG%"=="" set "CONFIG=pioneer1.toml"
set "VENV=%~dp0.venv"
set "VPY=%VENV%\Scripts\python.exe"

if exist "%VENV%\esap_ready.txt" goto run

echo First run: setting up Python for eSAP...
set "BASEPY="
py -3 -c "import sys; sys.exit(0 if sys.version_info >= (3, 11) else 1)" >nul 2>&1
if not errorlevel 1 set "BASEPY=py -3"
if defined BASEPY goto havepy
python -c "import sys; sys.exit(0 if sys.version_info >= (3, 11) else 1)" >nul 2>&1
if not errorlevel 1 set "BASEPY=python"
if defined BASEPY goto havepy
echo.
echo Python 3.11 or newer was not found on this computer.
echo Install it from https://www.python.org/downloads/ and tick "Add python.exe to PATH",
echo or open a command prompt and run:   winget install Python.Python.3.13
echo Then run this file again.
goto fail

:havepy
if exist "%VENV%" rmdir /s /q "%VENV%"
%BASEPY% -m venv "%VENV%"
if errorlevel 1 goto fail
"%VPY%" -m pip install --disable-pip-version-check --quiet -r requirements.txt
if errorlevel 1 goto fail
echo ready> "%VENV%\esap_ready.txt"
echo Python is set up.
echo.

:run
"%VPY%" esap.py --config "%CONFIG%" %2
set "RC=%ERRORLEVEL%"
goto end

:fail
echo.
echo Setup did not finish, so nothing was built. Check the internet connection and try again.
set "RC=9"

:end
rem Keep the window open when the file was double-clicked, so the result can be read.
echo %CMDCMDLINE% | find /i "%~nx0" >nul && pause
exit /b %RC%

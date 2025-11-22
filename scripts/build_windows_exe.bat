@echo off
setlocal enabledelayedexpansion

rem Switch to the repository root (one level above this script)
set SCRIPT_DIR=%~dp0
pushd "%SCRIPT_DIR%.." || exit /b 1

if not exist .venv (
    python -m venv .venv
    if errorlevel 1 goto :error
)

call .venv\Scripts\activate.bat
if errorlevel 1 goto :error

python -m pip install --upgrade pip
if errorlevel 1 goto :error

python -m pip install -r requirements.txt
if errorlevel 1 goto :error

python -m pip install pyinstaller
if errorlevel 1 goto :error

pyinstaller --noconfirm --clean bomcheck.spec
if errorlevel 1 goto :error

echo.
echo Build finished. The executable is located at dist\bomcheck\bomcheck.exe
goto :eof

:error
echo Build failed. Please check the errors above.
exit /b 1

endlocal

@echo off
setlocal enabledelayedexpansion

if not exist .venv (
    python -m venv .venv
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

pyinstaller --noconfirm --clean bomcheck.spec

echo.
echo Build finished. The executable is located at dist\bomcheck\bomcheck.exe
endlocal

@echo off
:: Build script for Windows — run this on a Windows machine
:: Requires: Python 3.10+, pip install -r requirements.txt pyinstaller

echo === Gelatik Automation Build ===
echo.

:: Install dependencies
pip install -r requirements.txt pyinstaller

:: Clean previous build
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build

:: Build executable
pyinstaller Automation.spec --clean

echo.
if exist dist\Automation.exe (
    echo BUILD SUCCESS: dist\Automation.exe
) else (
    echo BUILD FAILED
    exit /b 1
)

:: Create release folder
set RELEASE=dist\SOS Dashboard Automation
mkdir "%RELEASE%"
copy dist\Automation.exe "%RELEASE%\"
copy config.json "%RELEASE%\"
copy README.md "%RELEASE%\"
mkdir "%RELEASE%\logs"

echo.
echo Release folder: %RELEASE%
echo Done.

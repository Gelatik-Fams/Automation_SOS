@echo off
:: Build script for Windows — run this on a Windows machine
:: Requires: Python 3.10+, pip install -r requirements.txt pyinstaller

echo === Gelatik Automation Build ===
echo.

tasklist /FI "IMAGENAME eq Automation.exe" | %SystemRoot%\System32\findstr.exe /I "Automation.exe" >nul
if not errorlevel 1 (
    echo ERROR: Automation.exe masih berjalan.
    echo Tutup aplikasi Automation dulu, lalu jalankan build_windows.bat lagi.
    exit /b 1
)

:: Install dependencies
pip install -r requirements.txt pyinstaller

:: Clean previous build
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build

:: Build executable
python -m PyInstaller Automation.spec --clean

echo.
set "RELEASE=dist\SOS Dashboard Automation"
if exist "%RELEASE%\Automation.exe" (
    echo BUILD SUCCESS: %RELEASE%\Automation.exe
) else (
    echo BUILD FAILED
    exit /b 1
)

:: Add release files directly to the PyInstaller output folder
copy config.json "%RELEASE%\"
copy README.md "%RELEASE%\"
if not exist "%RELEASE%\logs" mkdir "%RELEASE%\logs"

echo.
echo Release folder: %RELEASE%
echo Done.

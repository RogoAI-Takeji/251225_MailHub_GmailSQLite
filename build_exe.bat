@echo off
REM ==========================================
REM RogoAI Mail Hub - EXE Build Script (JP Only)
REM ==========================================

echo.
echo ==========================================
echo RogoAI Mail Hub - EXE Builder (Japanese)
echo ==========================================
echo.

REM Check PyInstaller installation
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo [ERROR] PyInstaller is not installed.
    echo Please run: pip install pyinstaller
    echo.
    pause
    exit /b 1
)

echo [INFO] PyInstaller is installed.
echo.

REM Cleanup old builds
if exist dist (
    echo [INFO] Cleaning up old dist folder...
    rmdir /s /q dist
)

if exist build (
    echo [INFO] Cleaning up old build folder...
    rmdir /s /q build
)

echo.
echo ==========================================
echo Building Japanese Version
echo ==========================================
echo.

REM Build Japanese version
pyinstaller --onefile --windowed --name MailHub_v1 --icon=icon.ico --add-data "lib;lib" --hidden-import tkinterweb main.py

if errorlevel 1 (
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

echo.
echo [SUCCESS] Build completed successfully.
echo.

REM Create ZIP file
echo ==========================================
echo Creating ZIP file
echo ==========================================
echo.

cd dist

if exist MailHub_v1.zip del MailHub_v1.zip

REM Use PowerShell to create ZIP file
powershell -command "Compress-Archive -Path 'MailHub_v1.exe' -DestinationPath 'MailHub_v1.zip'"

cd ..

echo.
echo ==========================================
echo Build Complete!
echo ==========================================
echo.
echo Output files:
echo   - dist\MailHub_v1.exe
echo   - dist\MailHub_v1.zip
echo.
echo These files are ready to be uploaded to GitHub Releases.
echo.
echo Note: This version is Japanese only.
echo English version will be available in future releases.
echo.

pause
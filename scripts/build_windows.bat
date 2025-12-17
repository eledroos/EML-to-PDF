@echo off
REM Build script for Windows
setlocal enabledelayedexpansion

REM Get version
if exist VERSION (
    set /p VERSION=<VERSION
) else (
    set VERSION=2.0.0
)

set APP_NAME=EML-to-PDF

echo Building %APP_NAME% v%VERSION% for Windows...

REM Clean previous builds
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

REM Install dependencies
pip install -r requirements-dev.txt

REM Run build
python build.py --platform windows --package

echo.
echo Build complete!
echo Output: dist\%APP_NAME%-v%VERSION%-Windows.exe

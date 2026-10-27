@echo off
setlocal EnableExtensions
cd /d "%~dp0"
set "LANGUAGE=%~1"
if "%LANGUAGE%"=="" set "LANGUAGE=pl"
if /I not "%LANGUAGE%"=="pl" if /I not "%LANGUAGE%"=="en" (
    echo Usage: build_onedir.cmd [pl^|en]
    exit /b 2
)
if not exist ".venv\Scripts\python.exe" (
    echo ERROR: missing .venv - run: python -m venv .venv
    exit /b 1
)
set "ENTRY=launcher_pl.py"
set "NAME=XlsxDiff_GUI_PL"
if /I "%LANGUAGE%"=="en" (
    set "ENTRY=launcher_en.py"
    set "NAME=XlsxDiff_GUI_EN"
)
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onedir --windowed --name "%NAME%" --icon "assets\xlsxDiff.ico" --add-data "assets;assets" --add-data "LICENSE;." --add-data "NOTICE.txt;." --add-data "SOURCE.md;." --add-data "THIRD_PARTY.md;." "%ENTRY%"
if errorlevel 1 (
    echo ERROR: PyInstaller build failed.
    exit /b 1
)
echo Build successful: "%CD%\dist\%NAME%\%NAME%.exe"
echo Distribute entire output folder with same-version SOURCE ZIP, LICENSE, NOTICE.
endlocal

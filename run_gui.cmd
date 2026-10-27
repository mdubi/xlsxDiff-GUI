@echo off
cd /d "%~dp0"
if exist ".venv\Scripts\pythonw.exe" (
    start "xlsxDiff GUI" ".venv\Scripts\pythonw.exe" "xlsxDiff_GUI.py"
) else (
    echo Nie znaleziono .venv. Wykonaj instrukcje w README_PL.md
    pause
)

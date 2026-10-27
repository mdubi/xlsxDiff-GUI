@echo off
setlocal
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" package_source.py
) else (
    python package_source.py
)
if errorlevel 1 exit /b 1
echo Source package created. Inspect contents before uploading to GitHub.
endlocal

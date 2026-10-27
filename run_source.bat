@echo off
setlocal
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" source_launcher.py run
) else (
    echo Please run setup_source.bat first. See SOURCE_EDITION.md for instructions.
    pause
    exit /b 1
)
if errorlevel 1 (
    echo Startup failed. See source-launch.log. Keep this window open to read the error.
    pause
    exit /b 1
)
exit /b 0
exit /b 1

@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if not errorlevel 1 (
    py -3 source_launcher.py setup
) else (
    python source_launcher.py setup
)
if errorlevel 1 (
    echo Setup failed. Check source-launch.log and confirm that Python is installed.
) else (
    echo Setup complete. Next time, double-click run_source.bat.
)
pause

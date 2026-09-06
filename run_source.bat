@echo off
setlocal
cd /d "%~dp0"
py -m pip install -r requirements.txt
if errorlevel 1 goto :error
py file_mapping_tool.py
exit /b %errorlevel%

:error
echo.
echo Dependency installation failed. Check Python and network access.
pause
exit /b 1

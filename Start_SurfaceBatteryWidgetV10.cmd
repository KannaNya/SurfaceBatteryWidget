@echo off
setlocal

set "PYTHONW=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\pythonw.exe"
if exist "%PYTHONW%" goto run

set "PYTHONW=%LOCALAPPDATA%\Python\bin\pythonw.exe"
if exist "%PYTHONW%" goto run

for /f "delims=" %%I in ('where pythonw.exe 2^>nul') do if not defined PYTHONW set "PYTHONW=%%I"
if not defined PYTHONW (
    echo Python was not found. Install Python and run: python -m pip install -r requirements.txt
    pause
    exit /b 1
)

:run
start "" "%PYTHONW%" "%~dp0SurfaceBatteryWidgetV10.py"

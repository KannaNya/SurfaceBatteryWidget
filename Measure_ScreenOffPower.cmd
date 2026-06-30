@echo off
setlocal

set "PYTHON=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%PYTHON%" goto run

for /f "delims=" %%I in ('where python.exe 2^>nul') do if not defined PYTHON set "PYTHON=%%I"
if not defined PYTHON (
    echo Python was not found. Install Python and run: python -m pip install -r requirements.txt
    pause
    exit /b 1
)

:run
"%PYTHON%" "%~dp0MeasureScreenOffPower.py" --baseline-seconds 60 --screen-off-seconds 600
pause

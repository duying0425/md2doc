@echo off
setlocal

cd /d "%~dp0"
set "PYTHONPATH=%CD%\src;%PYTHONPATH%"

python -m md2doc
if errorlevel 1 (
    echo.
    echo Program exited with an error.
    pause
)

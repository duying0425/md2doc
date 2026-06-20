@echo off
setlocal

cd /d "%~dp0"
set "PYTHONPATH=%CD%\src;%PYTHONPATH%"

set "PYTHON_EXE=%CD%\.venv\Scripts\python.exe"
if exist "%PYTHON_EXE%" (
    "%PYTHON_EXE%" -m md2doc
    goto done
)

python -m md2doc

:done
if errorlevel 1 (
    echo.
    echo Program exited with an error.
    pause
)

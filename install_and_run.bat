@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title PROJECTZ GS1 DataMatrix 4.0.7.7 - Setup and Run

echo ============================================================
echo PROJECTZ GS1 DataMatrix 4.0.7.7
echo ============================================================
echo.

if not exist "app.py" (
    echo [ERROR] app.py was not found.
    pause
    exit /b 1
)

where py.exe >nul 2>nul
if not errorlevel 1 (
    set "PYTHON=py"
    goto PYTHON_FOUND
)

where python.exe >nul 2>nul
if not errorlevel 1 (
    set "PYTHON=python"
    goto PYTHON_FOUND
)

echo [ERROR] Python 3.11 or newer was not found.
echo Install Python and enable "Add Python to PATH".
pause
exit /b 1

:PYTHON_FOUND
if not exist "venv\Scripts\python.exe" (
    echo [1/3] Creating virtual environment...
    %PYTHON% -m venv venv
    if errorlevel 1 goto ERROR
)

echo [2/3] Installing/updating libraries...
"venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto ERROR
"venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto ERROR

where gswin64c.exe >nul 2>nul
if errorlevel 1 (
    where gswin32c.exe >nul 2>nul
)
if errorlevel 1 (
    echo.
    echo [WARNING] Ghostscript was not found in PATH.
    echo Install Ghostscript 64-bit, then restart this BAT.
    echo Official package: https://www.ghostscript.com/releases/gsdnld.html
    echo.
)

echo [3/3] Starting application...
"venv\Scripts\python.exe" app.py
if errorlevel 1 goto ERROR
exit /b 0

:ERROR
echo.
echo [ERROR] Operation failed.
pause
exit /b 1

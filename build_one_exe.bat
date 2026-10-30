@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title PROJECTZ GS1 DataMatrix 4.0.7.5 - Build EXE

echo ============================================================
echo BUILD ONE EXE
echo ============================================================
echo.

if not exist ".build311fresh\Scripts\python.exe" (
    echo Build environment not found: .build311fresh
    echo Create it with Python 3.11 and install requirements.txt first.
    pause
    exit /b 1
)


".build311fresh\Scripts\python.exe" -m PyInstaller ^
  --noconfirm --clean "PROJECTZ GS1 DataMatrix 4.0.7.5.spec"

if errorlevel 1 goto ERROR

echo.
echo EXE created:
echo %CD%\dist\PROJECTZ GS1 DataMatrix 4.0.7.5.exe
pause
exit /b 0

:ERROR
echo.
echo [ERROR] EXE build failed.
pause
exit /b 1

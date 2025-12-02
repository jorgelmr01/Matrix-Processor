@echo off
cd /d "%~dp0"
title Matrix Processor
color 0A

echo.
echo ========================================
echo   MATRIX PROCESSOR
echo ========================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    where python3 >nul 2>nul
    if errorlevel 1 (
        echo ERROR: Python is not installed!
        echo.
        echo Please install Python from: https://python.org/
        echo.
        pause
        goto :eof
    )
    python3 app.py
    goto :end
)

python app.py

:end
echo.
echo Application closed.
pause

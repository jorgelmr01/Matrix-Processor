@echo off
cd /d "%~dp0"
title Procesador de Matrices
color 0A

echo.
echo ========================================
echo   PROCESADOR DE MATRICES
echo ========================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    where python3 >nul 2>nul
    if errorlevel 1 (
        echo ERROR: Python no esta instalado!
        echo.
        echo Por favor instala Python desde: https://python.org/
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
echo Aplicacion cerrada.
pause

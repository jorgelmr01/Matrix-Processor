@echo off
cd /d "%~dp0"
title Matrix Processor
color 0A

echo.
echo ========================================
echo   MATRIX PROCESSOR
echo ========================================
echo.

where node >nul 2>nul
if errorlevel 1 (
    echo ERROR: Node.js is not installed!
    echo.
    echo Please download and install Node.js from:
    echo https://nodejs.org/
    echo.
    echo After installing, run this file again.
    echo.
    pause
    goto :eof
)

echo Checking dependencies...
if not exist "node_modules\" (
    echo.
    echo Installing dependencies for the first time...
    echo This may take a few minutes. Please wait...
    echo.
    npm install
    if errorlevel 1 (
        echo.
        echo ERROR: Failed to install dependencies!
        echo Check your internet connection and try again.
        echo.
        pause
        goto :eof
    )
    echo.
    echo Dependencies installed!
    echo.
)

echo.
echo Starting application...
echo.
echo The app will open in your browser at: http://localhost:5173
echo.
echo Keep this window open while using the app.
echo Close this window to stop the app.
echo.
echo ========================================
echo.

start http://localhost:5173
npm run dev

echo.
echo Application closed.
pause

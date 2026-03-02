@echo off
title EDP App - Setup and Launch
color 1F

echo.
echo  ============================================
echo   EDP 2026 - Professional Development App
echo   Windows Setup Script
echo  ============================================
echo.

:: Check if Node.js is installed
node --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo  [!] Node.js is NOT installed yet.
    echo.
    echo  Please do the following:
    echo  1. Open your browser
    echo  2. Go to: https://nodejs.org
    echo  3. Click the green LTS button to download
    echo  4. Run the installer, click Next until done
    echo  5. RESTART your computer
    echo  6. Then double-click START-APP.bat again
    echo.
    pause
    exit /b
)

echo  [OK] Node.js is installed:
node --version
echo.

:: Install packages if not already installed
IF NOT EXIST "node_modules" (
    echo  [>>] Installing packages for the first time...
    echo      This takes 2-3 minutes. Please wait.
    echo.
    call npm install --legacy-peer-deps
    echo.
    echo  [OK] Packages installed!
    echo.
)

echo  Starting the EDP app...
echo.
echo  ============================================
echo   App will open in your browser shortly.
echo   Address: http://localhost:3000
echo.
echo   To STOP the app: press Ctrl+C
echo  ============================================
echo.

start "" "http://localhost:3000"
call npm start

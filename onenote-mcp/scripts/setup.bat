@echo off
setlocal enabledelayedexpansion

echo.
echo ===================================================
echo OneNote MCP Server: Bootstrap Setup (Windows)
echo ===================================================
echo.

:: Check Node.js
where node >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Node.js is not installed or not in your PATH.
    echo.
    echo Please download and install Node.js. v18 or higher is recommended.
    echo Get it from: https://nodejs.org/
    echo.
    echo After installing, restart your terminal and run this script again.
    pause
    exit /b 1
)

:: Check NPM
where npm >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] npm is not installed or not in your PATH.
    echo Please ensure npm is installed alongside Node.js.
    pause
    exit /b 1
)

:: If Node.js and NPM are present, delegate to setup.js
echo [INFO] Node.js and npm found. Starting main setup...
node "%~dp0setup.js"

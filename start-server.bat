@echo off
cd /d "%~dp0"

set "NODE_EXE="
set "NODE_DIR="
if exist "%ProgramFiles%\nodejs\node.exe" (
    set "NODE_EXE=%ProgramFiles%\nodejs\node.exe"
    set "NODE_DIR=%ProgramFiles%\nodejs"
)
if "%NODE_EXE%"=="" if exist "%ProgramFiles(x86)%\nodejs\node.exe" (
    set "NODE_EXE=%ProgramFiles(x86)%\nodejs\node.exe"
    set "NODE_DIR=%ProgramFiles(x86)%\nodejs"
)
if "%NODE_EXE%"=="" if exist "%LOCALAPPDATA%\Programs\node\node.exe" (
    set "NODE_EXE=%LOCALAPPDATA%\Programs\node\node.exe"
    set "NODE_DIR=%LOCALAPPDATA%\Programs\node"
)
if "%NODE_EXE%"=="" if exist "%USERPROFILE%\scoop\apps\nodejs\current\node.exe" (
    set "NODE_EXE=%USERPROFILE%\scoop\apps\nodejs\current\node.exe"
    set "NODE_DIR=%USERPROFILE%\scoop\apps\nodejs\current"
)

if "%NODE_EXE%"=="" (
    where node >nul 2>&1
    if %errorlevel% equ 0 (
        set "NODE_EXE=node"
        set "NODE_DIR="
    )
)

if "%NODE_EXE%"=="" (
    echo [ERROR] Node.js not found.
    echo Install from https://nodejs.org/ or run from Command Prompt.
    pause
    exit /b 1
)

if defined NODE_DIR set "PATH=%NODE_DIR%;%PATH%"

if not exist "node_modules" (
    echo Installing dependencies...
    call npm install
    echo.
)

echo Starting server...
"%NODE_EXE%" server.js

@echo off
REM Launcher for OneNote to Notion Import Script
REM Automatically uses PowerShell 7 (required for modern .NET HTTP APIs)

echo.
echo ========================================
echo   OneNote to Notion Import Launcher
echo ========================================
echo.

REM Check if PowerShell 7 is installed
where pwsh.exe >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: PowerShell 7 is not installed!
    echo.
    echo This script requires PowerShell 7.0 or higher.
    echo.
    echo To install PowerShell 7:
    echo   1. Download from: https://aka.ms/powershell
    echo   2. Or use winget: winget install Microsoft.PowerShell
    echo.
    pause
    exit /b 1
)

echo Starting import with PowerShell 7...
echo.

REM Run with PowerShell 7 (Core)
pwsh.exe -ExecutionPolicy Bypass -File "%~dp0Import-OneNoteToNotion.ps1"

echo.
echo ========================================
echo   Import Complete
echo ========================================
echo.
pause

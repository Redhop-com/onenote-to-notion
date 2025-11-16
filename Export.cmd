@echo off
REM Launcher for OneNote Export Script
REM Automatically uses Windows PowerShell 5.1 (required for OneNote COM automation)

echo.
echo ========================================
echo   OneNote to PDF Export Launcher
echo ========================================
echo.
echo Starting export with Windows PowerShell 5.1...
echo.

REM Run with Windows PowerShell 5.1 (NOT PowerShell Core)
powershell.exe -ExecutionPolicy Bypass -File "%~dp0Export-OneNoteToPDF.ps1"

echo.
echo ========================================
echo   Export Complete
echo ========================================
echo.
pause

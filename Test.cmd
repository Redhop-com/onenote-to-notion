@echo off
REM Launcher for OneNote Connection Test Script
REM Automatically uses Windows PowerShell 5.1 (required for OneNote COM automation)

echo.
echo ========================================
echo   OneNote Connection Test Launcher
echo ========================================
echo.
echo Testing OneNote connection with Windows PowerShell 5.1...
echo.

REM Run with Windows PowerShell 5.1 (NOT PowerShell Core)
powershell.exe -ExecutionPolicy Bypass -File "%~dp0Test-OneNoteConnection.ps1"

echo.
echo ========================================
echo   Test Complete
echo ========================================
echo.
pause

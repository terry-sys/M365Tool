@echo off
setlocal
chcp 65001 >nul

cd /d "%~dp0"

echo ========================================
echo M365Tool - Run Program
echo ========================================
echo.

if not exist "M365Tool.csproj" (
    echo Project file not found: M365Tool.csproj
    pause
    exit /b 1
)

where dotnet >nul 2>nul
if %errorlevel% neq 0 (
    echo dotnet was not found. Please install the .NET SDK first.
    pause
    exit /b 1
)

echo Building project...
dotnet build "M365Tool.csproj" --configuration Release

if %errorlevel% neq 0 (
    echo Build failed.
    echo Please review the error output and fix the issue.
    pause
    exit /b 1
)

set "EXE_PATH=bin\Release\net8.0-windows7.0\M365Tool.exe"

if not exist "%EXE_PATH%" (
    echo Build succeeded, but the executable was not found:
    echo %EXE_PATH%
    pause
    exit /b 1
)

echo.
echo Build succeeded.
echo.
echo Starting program with administrator privileges...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process '%EXE_PATH%' -Verb RunAs"

echo.
echo Program started.
pause


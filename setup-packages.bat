@echo off
setlocal enabledelayedexpansion

echo === NuGet Package Setup ===

:: packages.config ファイルの存在確認
if not exist "packages.config" (
    echo ERROR: packages.config file not found!
    echo Please ensure packages.config exists in the current directory.
    pause
    exit /b 1
)

echo Reading packages from packages.config...

:: PowerShell 実行ポリシーをチェック
powershell -Command "Get-ExecutionPolicy" | findstr /C:"Unrestricted" >nul
if %ERRORLEVEL% neq 0 (
    echo Warning: PowerShell execution policy may prevent script execution.
    echo If the script fails, run this command as administrator:
    echo   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
    echo.
)

:: PowerShell スクリプトで packages.config を処理
echo Running PowerShell setup script for all packages...
powershell -ExecutionPolicy Bypass -File "%~dp0setup-packages.ps1"

if %ERRORLEVEL% equ 0 (
    echo.
    echo ========================================
    echo All packages setup completed successfully.
    echo ========================================
) else (
    echo.
    echo ========================================
    echo Setup failed. Please check the error messages above.
    echo ========================================
)

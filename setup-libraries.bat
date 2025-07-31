@echo off
echo === NuGet Package Setup ===

:: PowerShell 実行ポリシーをチェック
powershell -Command "Get-ExecutionPolicy" | findstr /C:"Unrestricted" >nul
if %ERRORLEVEL% neq 0 (
    echo Warning: PowerShell execution policy may prevent script execution.
    echo If the script fails, run this command as administrator:
    echo   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
    echo.
    pause
)

:: PowerShell スクリプトの実行
echo Running PowerShell setup script...
:: TODO: 複数パッケージのスクリプト構成やエラーハンドリングが未調整
powershell -ExecutionPolicy Bypass -File "%~dp0setup-libraries.ps1" -PackageName "DocumentFormat.OpenXml" 
powershell -ExecutionPolicy Bypass -File "%~dp0setup-libraries.ps1" -PackageName "DocumentFormat.OpenXml.Framework" 
powershell -ExecutionPolicy Bypass -File "%~dp0setup-libraries.ps1" -PackageName "System.IO.Packaging" 

:: tools 以下にものができる
powershell -ExecutionPolicy Bypass -File "%~dp0setup-libraries.ps1" -PackageName "Microsoft.Net.Compilers" 

if %ERRORLEVEL% equ 0 (
    echo.
    echo Setup completed successfully!
) else (
    echo.
    echo Setup failed. Please check the error messages above.
)

pause

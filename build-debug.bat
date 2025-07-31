@echo off
setlocal enabledelayedexpansion

echo === C# Debug Build Script ===

:: csc.exeのパスを動的に解決
echo Searching for csc.exe...
for /f "tokens=*" %%i in ('where csc 2^>nul') do (
    set CSC_PATH=%%i
    goto :found_csc
)

:: where で見つからない場合は標準的なパスを検索
echo csc.exe not found in PATH, searching standard locations...
set "POSSIBLE_PATHS=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
set "POSSIBLE_PATHS=!POSSIBLE_PATHS! C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"

for %%p in (!POSSIBLE_PATHS!) do (
    if exist "%%p" (
        set CSC_PATH=%%p
        goto :found_csc
    )
)

echo ERROR: csc.exe not found!
echo Please ensure .NET Framework is installed.
exit /b 1

:found_csc
echo Found csc.exe at: !CSC_PATH!

:: 出力ディレクトリの作成
if not exist "bin" mkdir bin

:: DLLの存在確認
if not exist "lib\DocumentFormat.OpenXml.dll" (
    echo WARNING: DocumentFormat.OpenXml.dll not found in lib directory
)

echo Building with debug information...

:: FIXME:
set CSC_PATH=packages\Microsoft.Net.Compilers\tools\csc.exe

:: デバッグビルドの実行
"!CSC_PATH!" ^
    /target:exe ^
    /langversion:7 ^
    /debug+ ^
    /debug:full ^
    /optimize- ^
    /out:bin\%~n0.exe ^
    /reference:lib\DocumentFormat.OpenXml.dll ^
    /reference:lib\DocumentFormat.OpenXml.Framework.dll ^
    /reference:lib\System.IO.Packaging.dll ^
    /reference:System.dll ^
    /reference:System.Core.dll ^
    /reference:System.Xml.dll ^
    /reference:System.Xml.Linq.dll ^
    /reference:WPF\WindowsBase.dll ^
    src\*.cs

if %ERRORLEVEL% equ 0 (
    echo Build completed successfully!
    echo Copying required DLLs to bin directory...
    if exist "lib\DocumentFormat.OpenXml.dll" (
        copy "lib\DocumentFormat.OpenXml.dll" "bin\" >nul 2>&1
    )
    if exist "lib\DocumentFormat.OpenXml.Framework.dll" (
        copy "lib\DocumentFormat.OpenXml.Framework.dll" "bin\" >nul 2>&1
    )
    if exist "lib\System.IO.Packaging.dll" (
        copy "lib\System.IO.Packaging.dll" "bin\" >nul 2>&1
    )
    echo Executable: bin\%~n0.exe
) else (
    echo Build failed with error code %ERRORLEVEL%
    exit /b %ERRORLEVEL%
)

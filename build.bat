@echo off
setlocal enabledelayedexpansion

:: パラメータの解析
set BUILD_TYPE=debug
set SHOW_HELP=false

:parse_args
if "%~1"=="" goto :args_parsed
if /i "%~1"=="debug" (
    set BUILD_TYPE=debug
    shift
    goto :parse_args
)
if /i "%~1"=="release" (
    set BUILD_TYPE=release
    shift
    goto :parse_args
)
if /i "%~1"=="-h" (
    set SHOW_HELP=true
    shift
    goto :parse_args
)
if /i "%~1"=="--help" (
    set SHOW_HELP=true
    shift
    goto :parse_args
)
if /i "%~1"=="help" (
    set SHOW_HELP=true
    shift
    goto :parse_args
)
echo Unknown parameter: %~1
echo Use 'build help' for usage information.
exit /b 1

:args_parsed

:: ヘルプの表示
if "%SHOW_HELP%"=="true" (
    echo === C# Build Script ===
    echo.
    echo Usage: build [BUILD_TYPE] [OPTIONS]
    echo.
    echo BUILD_TYPE:
    echo   debug      Build with debug information ^(default^)
    echo   release    Build optimized release version
    echo.
    echo OPTIONS:
    echo   -h, --help Show this help message
    echo.
    echo Examples:
    echo   build           ^(builds debug version^)
    echo   build debug     ^(builds debug version^)
    echo   build release   ^(builds release version^)
    echo.
    exit /b 0
)

echo === C# %BUILD_TYPE% Build Script ===

:: カスタムコンパイラパスの設定 (コメントのFIXMEを反映)
if exist "packages\Microsoft.Net.Compilers\tools\csc.exe" (
    set CSC_PATH=packages\Microsoft.Net.Compilers\tools\csc.exe
    goto :found_csc
)

echo ERROR: csc.exe not found!
echo Please run setup-libraries.bat first
exit /b 1

:found_csc
echo Found csc.exe at: !CSC_PATH!

:: 出力ディレクトリの作成
if not exist "bin" mkdir bin
if not exist "bin\debug" mkdir bin\debug
if not exist "bin\release" mkdir bin\release

:: ビルドタイプに応じた設定
if /i "%BUILD_TYPE%"=="debug" (
    echo Building with debug information...
    set "DEBUG_FLAGS=/debug+ /debug:portable /optimize-"
    set "OUTPUT_SUFFIX=debug"
) else (
    echo Building optimized release version...
    set "DEBUG_FLAGS=/optimize+"
    set "OUTPUT_SUFFIX=release"
)

:: 実行ファイル名の設定
set "OUTPUT_DIR=bin\%OUTPUT_SUFFIX%"
set "OUTPUT_NAME=%OUTPUT_DIR%\App.exe"

:: ビルドの実行
"!CSC_PATH!" ^
    /target:exe ^
    /langversion:7 ^
    !DEBUG_FLAGS! ^
    /out:!OUTPUT_NAME! ^
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
    echo Build completed successfully.
    echo Copying required DLLs to %OUTPUT_DIR% directory...
    
    :: 必要な DLL を出力ディレクトリにコピー
    copy "lib\*.dll" "%OUTPUT_DIR%\" >nul 2>&1
    
    echo Executable: !OUTPUT_NAME!
    
    :: 実行方法を表示
    echo.
    echo To run the application:
    echo   !OUTPUT_NAME!
    
) else (
    echo Build failed with error code %ERRORLEVEL%
    exit /b %ERRORLEVEL%
)

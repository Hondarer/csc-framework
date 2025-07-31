# pure-csc-framework

Windows 組み込みの csc.exe と VSCode を使用し、DocumentFormat.OpenXml ライブラリで Excel ファイル操作を行う C# 開発環境の構築手順です。

## 概要

このガイドでは、以下の制約条件下でC#Excel操作アプリケーションを開発します。

### 制約条件

- **追加インストール禁止**: Visual Studio、.NET SDK等のインストール不可
- **Windows標準機能のみ**: 組み込みのcsc.exeとVSCodeプラグインのみ使用
- **外部ライブラリ利用**: NuGetパッケージの手動取得と配置

### 実現する機能

- 本格的な Excel (.xlsx) ファイルの読み書き
- 複数シート対応
- フルデバッグ環境 (ブレークポイント、ステップ実行)
- 自動ビルドシステム
- プロジェクト構造の標準化

## 前提条件と環境確認

### .NET Framework バージョンの確認

コマンドプロンプトで以下を実行します。

```cmd
dir C:\Windows\Microsoft.NET\Framework64\
```

**必要条件:** v4.0.30319 フォルダが存在すること

### csc.exe の確認

コマンドプロンプトで以下を実行します。

```cmd
dir C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
```

**期待される出力例:**

```txt
 C:\Windows\Microsoft.NET\Framework64\v4.0.30319 のディレクトリ

2024/11/08  07:53         2,569,696 csc.exe
               1 個のファイル           2,569,696 バイト
               0 個のディレクトリ  1,899,705,847,808 バイトの空き領域
```

## 開発環境のセットアップ

### VSCodeプラグインのインストール

拡張機能タブ (`Ctrl+Shift+X`) から以下をインストール：

**必須プラグイン:**

- `C#` (Microsoft製) - 基本的な C# サポート
- `C# Dev Kit` (Microsoft製) - 新しい C# 開発体験

**推奨プラグイン:**

- `Error Lens` - エラーの可視化
- `Bracket Pair Colorizer` - 括弧の色分け

### 基本フォルダ構造の作成

```cmd
mkdir bin
mkdir src
mkdir .vscode
```

**完成後の構造:**

```txt
ExcelApp/
├── .vscode/              # VSCode設定
├── bin/                  # ビルド出力
├── lib/                  # 外部ライブラリDLL
├── packages/             # NuGetパッケージ展開場所
├── src/                  # ソースコード
├── build-debug.bat       # デバッグビルドスクリプト
├── build-release.bat     # リリースビルドスクリプト
├── clean.bat             # クリーンスクリプト
├── setup-libraries.ps1   # PowerShell自動セットアップ
├── setup-libraries.bat   # バッチファイル版セットアップ
└── setup-project.bat     # 統合セットアップ
```

## NuGet ライブラリの取得

### PowerShell 自動セットアップスクリプト

#### `setup-libraries.ps1` の作成

```powershell:setup-libraries.ps1
# setup-libraries.ps1
# NuGetパッケージの自動ダウンロードと展開スクリプト

param(
    [string]$PackageName = "DocumentFormat.OpenXml",
    [string]$Version = "latest",
    [string]$TargetFramework = "net462"
)

# 設定
$PackagesDir = "packages"
$LibDir = "lib"
$TempDir = "temp"

Write-Host "=== NuGet Package Auto Setup ===" -ForegroundColor Green
Write-Host "Package: $PackageName" -ForegroundColor Cyan
Write-Host "Version: $Version" -ForegroundColor Cyan
Write-Host "Target Framework: $TargetFramework" -ForegroundColor Cyan

# 必要なディレクトリの作成
@($PackagesDir, $LibDir, $TempDir) | ForEach-Object {
    if (!(Test-Path $_)) {
        New-Item -ItemType Directory -Path $_ -Force | Out-Null
        Write-Host "Created directory: $_" -ForegroundColor Yellow
    }
}

# NuGet APIからパッケージ情報を取得
function Get-LatestPackageVersion {
    param([string]$PackageName)
    
    try {
        Write-Host "Fetching package information from NuGet API..." -ForegroundColor Blue
        $apiUrl = "https://api.nuget.org/v3-flatcontainer/$($PackageName.ToLower())/index.json"
        $response = Invoke-RestMethod -Uri $apiUrl -ErrorAction Stop
        $latestVersion = $response.versions | Select-Object -Last 1
        return $latestVersion
    }
    catch {
        Write-Warning "Failed to fetch version info from API: $($_.Exception.Message)"
        return $null
    }
}

# パッケージバージョンの決定
if ($Version -eq "latest") {
    $actualVersion = Get-LatestPackageVersion -PackageName $PackageName
    if ($actualVersion) {
        Write-Host "Latest version found: $actualVersion" -ForegroundColor Green
        $Version = $actualVersion
    } else {
        Write-Host "Using fallback version: 3.0.1" -ForegroundColor Yellow
        $Version = "3.0.1"
    }
} else {
    Write-Host "Using specified version: $Version" -ForegroundColor Green
}

# パッケージのダウンロード
$packageFileName = "$PackageName.$Version.nupkg"
$downloadUrl = "https://api.nuget.org/v3-flatcontainer/$($PackageName.ToLower())/$Version/$($PackageName.ToLower()).$Version.nupkg"
$downloadPath = Join-Path $TempDir $packageFileName

Write-Host "Downloading package from: $downloadUrl" -ForegroundColor Blue

try {
    $webClient = New-Object System.Net.WebClient

    # プログレス表示
    Register-ObjectEvent -InputObject $webClient -EventName DownloadProgressChanged -Action {
        Write-Progress -Activity "Downloading $using:packageFileName" -Status "$($EventArgs.ProgressPercentage)% Complete" -PercentComplete $EventArgs.ProgressPercentage
    }

    # 完了処理
    Register-ObjectEvent -InputObject $webClient -EventName DownloadFileCompleted -Action {
        Write-Progress -Activity "Downloading $using:packageFileName" -Completed
        if ($EventArgs.Error) {
            throw $EventArgs.Error
        }
        Write-Host "Download completed: $using:downloadPath" -ForegroundColor Green
    }

    $webClient.DownloadFileAsync([Uri]$downloadUrl, $downloadPath)
    
    # ダウンロード完了まで待機
    while ($webClient.IsBusy) {
        Start-Sleep -Milliseconds 100
    }
    
    $webClient.Dispose()
}
catch {
    Write-Error "Failed to download package: $($_.Exception.Message)"
    exit 1
}

# パッケージファイルの存在確認
if (!(Test-Path $downloadPath)) {
    Write-Error "Downloaded package file not found: $downloadPath"
    exit 1
}

Write-Host "Package downloaded successfully: $downloadPath" -ForegroundColor Green
Write-Host "File size: $([math]::Round((Get-Item $downloadPath).Length / 1MB, 2)) MB" -ForegroundColor Cyan

# パッケージの展開
$extractPath = Join-Path $PackagesDir $PackageName

Write-Host "Extracting package to: $extractPath" -ForegroundColor Blue

try {
    # 既存の展開フォルダを削除
    if (Test-Path $extractPath) {
        Remove-Item $extractPath -Recurse -Force
        Write-Host "Removed existing extraction directory" -ForegroundColor Yellow
    }
    
    # ZIP として展開（.nupkg は ZIP ファイル）
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($downloadPath, $extractPath)
    
    Write-Host "Package extracted successfully" -ForegroundColor Green
}
catch {
    Write-Error "Failed to extract package: $($_.Exception.Message)"
    exit 1
}

# DLLファイルの検索とコピー
Write-Host "Searching for DLL files..." -ForegroundColor Blue

$libSearchPaths = @(
    (Join-Path $extractPath "lib\$TargetFramework"),
    (Join-Path $extractPath "lib\net48"),
    (Join-Path $extractPath "lib\net472"),
    (Join-Path $extractPath "lib\net471"),
    (Join-Path $extractPath "lib\net47"),
    (Join-Path $extractPath "lib\net46"),
    (Join-Path $extractPath "lib\netstandard2.0"),
    (Join-Path $extractPath "lib\netstandard2.1")
)

$dllsCopied = 0
$primaryDll = "$PackageName.dll"

foreach ($searchPath in $libSearchPaths) {
    if (Test-Path $searchPath) {
        Write-Host "Checking: $searchPath" -ForegroundColor Cyan
        
        $dllFiles = Get-ChildItem -Path $searchPath -Filter "*.dll" -File
        
        if ($dllFiles.Count -gt 0) {
            Write-Host "Found $($dllFiles.Count) DLL file(s) in $searchPath" -ForegroundColor Green
            
            foreach ($dll in $dllFiles) {
                $destPath = Join-Path $LibDir $dll.Name
                Copy-Item $dll.FullName $destPath -Force
                Write-Host "  Copied: $($dll.Name) -> lib\" -ForegroundColor Green
                $dllsCopied++
            }
            
            # 主要なDLLが見つかったら他のパスは検索しない
            if ($dllFiles.Name -contains $primaryDll) {
                Write-Host "Primary DLL found, skipping other target frameworks" -ForegroundColor Yellow
                break
            }
        }
    }
}

if ($dllsCopied -eq 0) {
    Write-Warning "No DLL files were found or copied!"
    Write-Host "Available lib directories:" -ForegroundColor Yellow
    
    $libBasePath = Join-Path $extractPath "lib"
    if (Test-Path $libBasePath) {
        Get-ChildItem $libBasePath -Directory | ForEach-Object {
            Write-Host "  - lib\$($_.Name)" -ForegroundColor Cyan
        }
    }
} else {
    Write-Host "Successfully copied $dllsCopied DLL file(s) to lib directory" -ForegroundColor Green
}

# 依存関係の確認
$nuspecPath = Join-Path $extractPath "$PackageName.nuspec"
if (Test-Path $nuspecPath) {
    Write-Host "Checking dependencies..." -ForegroundColor Blue
    
    try {
        [xml]$nuspec = Get-Content $nuspecPath
        $dependencies = $nuspec.package.metadata.dependencies.dependency
        
        if ($dependencies) {
            Write-Host "Dependencies found:" -ForegroundColor Yellow
            $dependencies | ForEach-Object {
                if ($_.id -and $_.version) {
                    Write-Host "  - $($_.id) (>= $($_.version))" -ForegroundColor Cyan
                }
            }
            Write-Host "Note: Dependencies are not automatically downloaded. Install manually if needed." -ForegroundColor Yellow
        } else {
            Write-Host "No dependencies found" -ForegroundColor Green
        }
    }
    catch {
        Write-Warning "Failed to parse .nuspec file: $($_.Exception.Message)"
    }
}

# 一時ファイルのクリーンアップ
Write-Host "Cleaning up temporary files..." -ForegroundColor Blue
Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue

# 結果の表示
Write-Host "`n=== Setup Summary ===" -ForegroundColor Green
Write-Host "Package: $PackageName v$Version" -ForegroundColor White
Write-Host "DLLs copied: $dllsCopied" -ForegroundColor White
Write-Host "Target directory: $LibDir" -ForegroundColor White

if (Test-Path (Join-Path $LibDir "$PackageName.dll")) {
    Write-Host "Setup completed successfully!" -ForegroundColor Green
    Write-Host "You can now build your project." -ForegroundColor Green
} else {
    Write-Warning "Setup may not have completed successfully."
    Write-Host "Please check the lib directory manually." -ForegroundColor Yellow
}

Write-Host "`nPress any key to continue..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
```

#### バッチファイルから PowerShell を実行

#### `setup-libraries.bat` の作成

```bat:setup-libraries.bat
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
powershell -ExecutionPolicy Bypass -File "%~dp0setup-libraries.ps1"

if %ERRORLEVEL% equ 0 (
    echo.
    echo Setup completed successfully!
) else (
    echo.
    echo Setup failed. Please check the error messages above.
)

pause
```

### 使用方法

**最も簡単な実行方法：**

```cmd
setup-libraries.bat
```

**PowerShell直接実行：**

```powershell
.\setup-libraries.ps1
```

**特定バージョン指定：**

```powershell
.\setup-libraries.ps1 -PackageName "DocumentFormat.OpenXml" -Version "2.20.0"
```

## ビルドシステムの構築

### 動的パス解決によるビルドスクリプト

#### `build-debug.bat` の作成

```bat:build-debug.bat
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

:: デバッグビルドの実行
"!CSC_PATH!" ^
    /target:exe ^
    /debug+ ^
    /debug:full ^
    /optimize- ^
    /out:bin\%~n0.exe ^
    /reference:lib\DocumentFormat.OpenXml.dll ^
    /reference:System.dll ^
    /reference:System.Core.dll ^
    /reference:System.IO.Packaging.dll ^
    /reference:System.Xml.dll ^
    /reference:System.Xml.Linq.dll ^
    /reference:WindowsBase.dll ^
    src\*.cs

if %ERRORLEVEL% equ 0 (
    echo Build completed successfully!
    echo Copying required DLLs to bin directory...
    if exist "lib\DocumentFormat.OpenXml.dll" (
        copy "lib\DocumentFormat.OpenXml.dll" "bin\" >nul 2>&1
    )
    echo Executable: bin\%~n0.exe
) else (
    echo Build failed with error code %ERRORLEVEL%
    exit /b %ERRORLEVEL%
)
```

#### `build-release.bat` の作成

```bat:build-release.bat
@echo off
setlocal enabledelayedexpansion

echo === C# Release Build Script ===

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

echo Building optimized release version...

:: リリースビルドの実行
"!CSC_PATH!" ^
    /target:exe ^
    /optimize+ ^
    /out:bin\%~n0.exe ^
    /reference:lib\DocumentFormat.OpenXml.dll ^
    /reference:System.dll ^
    /reference:System.Core.dll ^
    /reference:System.IO.Packaging.dll ^
    /reference:System.Xml.dll ^
    /reference:System.Xml.Linq.dll ^
    /reference:WindowsBase.dll ^
    src\*.cs

if %ERRORLEVEL% equ 0 (
    echo Build completed successfully!
    echo Copying required DLLs to bin directory...
    if exist "lib\DocumentFormat.OpenXml.dll" (
        copy "lib\DocumentFormat.OpenXml.dll" "bin\" >nul 2>&1
    )
    echo Executable: bin\%~n0.exe
) else (
    echo Build failed with error code %ERRORLEVEL%
    exit /b %ERRORLEVEL%
)
```

#### `clean.bat` の作成

```bat:clean.bat
@echo off
echo === Clean Build Output ===

if exist "bin" (
    echo Cleaning bin directory...
    del /Q "bin\*.*" 2>nul
    echo Build output cleaned.
) else (
    echo No bin directory found.
)

echo Clean completed.
```

### VSCode 統合設定

#### `.vscode/tasks.json` の作成

```json:tasks.json
{
    "version": "2.0.0",
    "tasks": [
        {
            "label": "build-debug",
            "type": "shell",
            "command": "${workspaceFolder}\\build-debug.bat",
            "group": {
                "kind": "build",
                "isDefault": true
            },
            "presentation": {
                "echo": true,
                "reveal": "always",
                "focus": false,
                "panel": "shared",
                "showReuseMessage": false
            },
            "options": {
                "cwd": "${workspaceFolder}"
            },
            "problemMatcher": "$msCompile"
        },
        {
            "label": "build-release",
            "type": "shell",
            "command": "${workspaceFolder}\\build-release.bat",
            "group": "build",
            "presentation": {
                "echo": true,
                "reveal": "always",
                "focus": false,
                "panel": "shared",
                "showReuseMessage": false
            },
            "options": {
                "cwd": "${workspaceFolder}"
            },
            "problemMatcher": "$msCompile"
        },
        {
            "label": "clean",
            "type": "shell",
            "command": "${workspaceFolder}\\clean.bat",
            "group": "build",
            "presentation": {
                "echo": true,
                "reveal": "always",
                "focus": false,
                "panel": "shared",
                "showReuseMessage": false
            },
            "options": {
                "cwd": "${workspaceFolder}"
            }
        }
    ]
}
```

---

## Excel操作プログラムの実装

### src/ExcelHandler.cs の作成

DocumentFormat.OpenXmlを使用したExcel操作クラス：

```csharp:ExcelHandler.cs
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelApp
{
    public class ExcelHandler
    {
        /// <summary>
        /// Excelファイルを読み込む
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        /// <param name="worksheetName">ワークシート名（nullの場合は最初のシート）</param>
        /// <returns>読み込んだデータ</returns>
        public static List<string[]> ReadExcel(string filePath, string worksheetName = null)
        {
            var result = new List<string[]>();
            
            try
            {
                using (var document = SpreadsheetDocument.Open(filePath, false))
                {
                    var workbookPart = document.WorkbookPart;
                    var worksheetPart = GetWorksheetPart(workbookPart, worksheetName);
                    
                    if (worksheetPart == null)
                    {
                        Console.WriteLine($"ワークシート '{worksheetName}' が見つかりません。");
                        return result;
                    }
                    
                    var worksheet = worksheetPart.Worksheet;
                    var sharedStringTablePart = workbookPart.SharedStringTablePart;
                    var sharedStringTable = sharedStringTablePart?.SharedStringTable;
                    
                    var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>();
                    
                    foreach (var row in rows.OrderBy(r => r.RowIndex))
                    {
                        var rowData = new List<string>();
                        var cells = row.Elements<Cell>().OrderBy(c => c.CellReference.Value);
                        
                        string lastColumnName = "";
                        foreach (var cell in cells)
                        {
                            var columnName = GetColumnName(cell.CellReference);
                            
                            // 空の列を埋める
                            while (GetNextColumnName(lastColumnName) != columnName && !string.IsNullOrEmpty(lastColumnName))
                            {
                                rowData.Add("");
                                lastColumnName = GetNextColumnName(lastColumnName);
                            }
                            
                            var cellValue = GetCellValue(cell, sharedStringTable);
                            rowData.Add(cellValue);
                            lastColumnName = columnName;
                        }
                        
                        result.Add(rowData.ToArray());
                    }
                }
                
                Console.WriteLine($"Excelファイルを読み込みました: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Excel読み込みエラー: {ex.Message}");
            }
            
            return result;
        }
        
        /// <summary>
        /// Excelファイルに書き込む
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        /// <param name="data">書き込むデータ</param>
        /// <param name="worksheetName">ワークシート名</param>
        public static void WriteExcel(string filePath, List<string[]> data, string worksheetName = "Sheet1")
        {
            try
            {
                // ファイルが存在する場合は削除
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
                
                using (var document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
                {
                    // Workbookパートを作成
                    var workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                    
                    // Worksheetパートを作成
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());
                    
                    // Sheetを追加
                    var sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
                    var sheet = new Sheet()
                    {
                        Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = worksheetName
                    };
                    sheets.Append(sheet);
                    
                    // データを書き込み
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                    
                    for (uint rowIndex = 0; rowIndex < data.Count; rowIndex++)
                    {
                        var row = new Row() { RowIndex = rowIndex + 1 };
                        sheetData.Append(row);
                        
                        for (int colIndex = 0; colIndex < data[(int)rowIndex].Length; colIndex++)
                        {
                            var cellReference = GetCellReference(rowIndex + 1, colIndex);
                            var cell = new Cell()
                            {
                                CellReference = cellReference,
                                DataType = CellValues.InlineString,
                                InlineString = new InlineString() { Text = new Text(data[(int)rowIndex][colIndex]) }
                            };
                            row.Append(cell);
                        }
                    }
                    
                    workbookPart.Workbook.Save();
                }
                
                Console.WriteLine($"Excelファイルを保存しました: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Excel書き込みエラー: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 複数シートでの書き込み
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        /// <param name="sheetsData">シート名とデータの辞書</param>
        public static void WriteMultipleSheets(string filePath, Dictionary<string, List<string[]>> sheetsData)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
                
                using (var document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
                {
                    var workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                    var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                    
                    uint sheetId = 1;
                    foreach (var sheetData in sheetsData)
                    {
                        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                        worksheetPart.Worksheet = new Worksheet(new SheetData());
                        
                        var sheet = new Sheet()
                        {
                            Id = workbookPart.GetIdOfPart(worksheetPart),
                            SheetId = sheetId++,
                            Name = sheetData.Key
                        };
                        sheets.Append(sheet);
                        
                        // データ書き込み
                        var sheetDataElement = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                        var data = sheetData.Value;
                        
                        for (uint rowIndex = 0; rowIndex < data.Count; rowIndex++)
                        {
                            var row = new Row() { RowIndex = rowIndex + 1 };
                            sheetDataElement.Append(row);
                            
                            for (int colIndex = 0; colIndex < data[(int)rowIndex].Length; colIndex++)
                            {
                                var cellReference = GetCellReference(rowIndex + 1, colIndex);
                                var cell = new Cell()
                                {
                                    CellReference = cellReference,
                                    DataType = CellValues.InlineString,
                                    InlineString = new InlineString() { Text = new Text(data[(int)rowIndex][colIndex]) }
                                };
                                row.Append(cell);
                            }
                        }
                    }
                    
                    workbookPart.Workbook.Save();
                }
                
                Console.WriteLine($"複数シートのExcelファイルを保存しました: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Excel書き込みエラー: {ex.Message}");
            }
        }
        
        #region ヘルパーメソッド
        
        private static WorksheetPart GetWorksheetPart(WorkbookPart workbookPart, string worksheetName)
        {
            if (string.IsNullOrEmpty(worksheetName))
            {
                return workbookPart.WorksheetParts.First();
            }
            
            var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>()
                .FirstOrDefault(s => s.Name == worksheetName);
            
            if (sheet == null) return null;
            
            return (WorksheetPart)workbookPart.GetPartById(sheet.Id);
        }
        
        private static string GetCellValue(Cell cell, SharedStringTable sharedStringTable)
        {
            if (cell.CellValue == null) return "";
            
            var value = cell.CellValue.InnerXml;
            
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                if (sharedStringTable != null)
                {
                    return sharedStringTable.ChildElements[int.Parse(value)].InnerText;
                }
            }
            
            return value;
        }
        
        private static string GetColumnName(string cellReference)
        {
            var columnName = "";
            foreach (char c in cellReference)
            {
                if (char.IsLetter(c))
                {
                    columnName += c;
                }
                else
                {
                    break;
                }
            }
            return columnName;
        }
        
        private static string GetNextColumnName(string columnName)
        {
            if (string.IsNullOrEmpty(columnName))
            {
                return "A";
            }
            
            var chars = columnName.ToCharArray();
            for (int i = chars.Length - 1; i >= 0; i--)
            {
                if (chars[i] == 'Z')
                {
                    chars[i] = 'A';
                }
                else
                {
                    chars[i]++;
                    return new string(chars);
                }
            }
            
            return "A" + new string(chars);
        }
        
        private static string GetCellReference(uint row, int column)
        {
            string columnName = "";
            int columnNumber = column;
            
            while (columnNumber >= 0)
            {
                columnName = (char)('A' + (columnNumber % 26)) + columnName;
                columnNumber = columnNumber / 26 - 1;
            }
            
            return columnName + row;
        }
        
        #endregion
    }
}
```

### 2. src/Program.cs の作成

メインプログラムの実装：

```csharp:Program.cs
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Excel ファイル読み書きサンプル - DocumentFormat.OpenXml使用");
            Console.WriteLine("==========================================================");
            
            // デバッグ用のブレークポイント設置箇所
            string excelFilePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.xlsx");
            
            try
            {
                // ステップ1: サンプルデータの作成
                Console.WriteLine("Step 1: サンプルデータを作成します...");
                var sampleData = CreateSampleData();
                
                // ステップ2: Excelファイルの書き込み
                Console.WriteLine("Step 2: Excelファイルに書き込みます...");
                ExcelHandler.WriteExcel(excelFilePath, sampleData, "社員リスト");
                
                // ステップ3: ファイルの存在確認
                Console.WriteLine("Step 3: ファイルの存在を確認します...");
                bool fileExists = File.Exists(excelFilePath);
                Console.WriteLine($"ファイル存在: {fileExists}");
                
                if (fileExists)
                {
                    // ステップ4: ファイル情報の表示
                    var fileInfo = new FileInfo(excelFilePath);
                    Console.WriteLine($"ファイルサイズ: {fileInfo.Length} bytes");
                    Console.WriteLine($"作成日時: {fileInfo.CreationTime}");
                    
                    // ステップ5: ファイルの読み込み
                    Console.WriteLine("Step 4: Excelファイルを読み込みます...");
                    var readData = ExcelHandler.ReadExcel(excelFilePath);
                    
                    // ステップ6: データの表示
                    Console.WriteLine("Step 5: データを表示します...");
                    DisplayData(readData);
                    
                    // ステップ7: データの処理
                    Console.WriteLine("Step 6: データを処理します...");
                    ProcessData(readData);
                    
                    // ステップ8: 複数シートのサンプル
                    Console.WriteLine("Step 7: 複数シートのファイルを作成します...");
                    CreateMultipleSheetSample();
                }
            }
            catch (Exception ex)
            {
                // エラー処理でもブレークポイントを設置可能
                Console.WriteLine($"エラーが発生しました: {ex.Message}");
                Console.WriteLine($"スタックトレース: {ex.StackTrace}");
            }
            
            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }
        
        /// <summary>
        /// サンプルデータを作成
        /// </summary>
        /// <returns>サンプルデータ</returns>
        static List<string[]> CreateSampleData()
        {
            var data = new List<string[]>();
            
            // ヘッダー行
            string[] headers = { "社員ID", "名前", "年齢", "部署", "給与", "入社日" };
            data.Add(headers);
            
            // データ行（ループでブレークポイント設置可能）
            string[][] employees = {
                new string[] { "EMP001", "田中太郎", "30", "開発部", "500000", "2020-04-01" },
                new string[] { "EMP002", "佐藤花子", "25", "デザイン部", "400000", "2021-07-15" },
                new string[] { "EMP003", "鈴木一郎", "35", "営業部", "600000", "2019-01-10" },
                new string[] { "EMP004", "高橋美咲", "28", "人事部", "450000", "2022-03-01" },
                new string[] { "EMP005", "山田和夫", "42", "管理部", "700000", "2018-08-20" }
            };
            
            foreach (var employee in employees)
            {
                // ここにブレークポイントを設置して各行の処理を確認
                data.Add(employee);
                Console.WriteLine($"従業員データを追加: {employee[1]} ({employee[3]})");
            }
            
            return data;
        }
        
        /// <summary>
        /// データを表示
        /// </summary>
        /// <param name="data">表示するデータ</param>
        static void DisplayData(List<string[]> data)
        {
            Console.WriteLine("\n=== データ表示 ===");
            
            for (int i = 0; i < data.Count; i++)
            {
                var row = data[i];
                Console.Write($"Row {i + 1}: ");
                
                for (int j = 0; j < row.Length; j++)
                {
                    // 各セルの値を確認できる
                    string cellValue = row[j];
                    Console.Write($"[{cellValue}] ");
                }
                Console.WriteLine();
            }
        }
        
        /// <summary>
        /// データを処理（給与統計の計算）
        /// </summary>
        /// <param name="data">処理するデータ</param>
        static void ProcessData(List<string[]> data)
        {
            if (data.Count <= 1) return; // ヘッダーのみの場合
            
            Console.WriteLine("\n=== データ処理（給与統計） ===");
            
            // 給与の合計を計算（デバッグで変数の値を確認）
            long totalSalary = 0;
            int employeeCount = 0;
            long maxSalary = 0;
            long minSalary = long.MaxValue;
            string highestPaidEmployee = "";
            
            for (int i = 1; i < data.Count; i++) // ヘッダーをスキップ
            {
                var row = data[i];
                if (row.Length > 4) // 給与列が存在するか確認
                {
                    if (long.TryParse(row[4], out long salary))
                    {
                        totalSalary += salary;
                        employeeCount++;
                        
                        // 最高給与と最低給与をチェック
                        if (salary > maxSalary)
                        {
                            maxSalary = salary;
                            highestPaidEmployee = row[1];
                        }
                        if (salary < minSalary)
                        {
                            minSalary = salary;
                        }
                        
                        // ここでsalary変数の値を確認できる
                        Console.WriteLine($"{row[1]}({row[3]}): {salary:N0}円");
                    }
                }
            }
            
            if (employeeCount > 0)
            {
                double averageSalary = (double)totalSalary / employeeCount;
                Console.WriteLine($"\n【統計結果】");
                Console.WriteLine($"従業員数: {employeeCount}人");
                Console.WriteLine($"給与合計: {totalSalary:N0}円");
                Console.WriteLine($"平均給与: {averageSalary:N0}円");
                Console.WriteLine($"最高給与: {maxSalary:N0}円 ({highestPaidEmployee})");
                Console.WriteLine($"最低給与: {minSalary:N0}円");
            }
        }
        
        /// <summary>
        /// 複数シートのサンプルファイルを作成
        /// </summary>
        static void CreateMultipleSheetSample()
        {
            var multiSheetData = new Dictionary<string, List<string[]>>();
            
            // 部署別売上データシート
            multiSheetData["部署別売上"] = new List<string[]>
            {
                new string[] { "部署", "Q1売上", "Q2売上", "Q3売上", "Q4売上", "年間合計" },
                new string[] { "開発部", "1200", "1350", "1100", "1450", "5100" },
                new string[] { "営業部", "2200", "2100", "2300", "2400", "9000" },
                new string[] { "デザイン部", "800", "900", "850", "950", "3500" },
                new string[] { "人事部", "400", "420", "380", "440", "1640" }
            };
            
            // 月別経費データシート
            multiSheetData["月別経費"] = new List<string[]>
            {
                new string[] { "月", "人件費", "オフィス費", "システム費", "その他", "合計" },
                new string[] { "1月", "3000", "500", "200", "300", "4000" },
                new string[] { "2月", "3100", "500", "250", "280", "4130" },
                new string[] { "3月", "3050", "520", "200", "350", "4120" }
            };
            
            string multiSheetFile = Path.Combine(Directory.GetCurrentDirectory(), "multi_sheet_sample.xlsx");
            ExcelHandler.WriteMultipleSheets(multiSheetFile, multiSheetData);
        }
    }
}
```

## デバッグ環境の設定

### `.vscode/launch.json` の作成

```json:launch.json
{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "Debug Launch",
            "type": "clr",
            "request": "launch",
            "program": "${workspaceFolder}/bin/build-debug.exe",
            "args": [],
            "cwd": "${workspaceFolder}",
            "stopAtEntry": false,
            "console": "integratedTerminal",
            "env": {
                "PATH": "${workspaceFolder}/lib;${env:PATH}"
            },
            "enableStepFiltering": true,
            "justMyCode": true,
            "preLaunchTask": "build-debug"
        },
        {
            "name": "Release Launch",
            "type": "clr",
            "request": "launch",
            "program": "${workspaceFolder}/bin/build-release.exe",
            "args": [],
            "cwd": "${workspaceFolder}",
            "stopAtEntry": false,
            "console": "externalTerminal",
            "env": {
                "PATH": "${workspaceFolder}/lib;${env:PATH}"
            },
            "preLaunchTask": "build-release"
        }
    ]
}
```

### デバッグ情報付きビルドの仕組み

**重要なコンパイラオプション:**

- `/debug+` : デバッグ情報を生成
- `/debug:full` : 完全なデバッグ情報
- `/optimize-` : 最適化を無効化（デバッグしやすくする）

これらのオプションにより、以下が可能になります：

- ブレークポイントの設置
- 変数値の確認
- ステップ実行
- コールスタックの表示

## ビルドと実行

### 初回セットアップ

**統合セットアップの実行:**

```cmd
setup-project.bat
```

このスクリプトにより以下が自動実行されます：

- 必要なディレクトリの作成
- NuGetパッケージの自動ダウンロード
- 基本ソースファイルのテンプレート生成
- セットアップ状況の検証

### 2. ビルド方法

**VSCodeでのビルド:**

- `Ctrl + Shift + B` → `build-debug` を選択

**コマンドラインでのビルド:**

```cmd
# デバッグビルド
build-debug.bat

# リリースビルド
build-release.bat

# クリーン
clean.bat
```

### 3. ビルド成功の確認

**期待される出力:**

```
=== C# Debug Build Script ===
Searching for csc.exe...
Found csc.exe at: C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
Building with debug information...
Build completed successfully!
Copying required DLLs to bin directory...
Executable: bin\build-debug.exe
```

### 4. 実行方法

**デバッグ実行:**

- `F5` キー（自動的にビルドしてから実行）

**デバッグなし実行:**

- `Ctrl + F5`

**コマンドライン実行:**

```cmd
bin\build-debug.exe
```

### 5. 出力ファイルの確認

実行成功後、以下のファイルが生成されます：

- `bin\build-debug.exe` - 実行可能ファイル
- `bin\build-debug.pdb` - デバッグ情報ファイル
- `bin\DocumentFormat.OpenXml.dll` - 依存ライブラリ
- `sample.xlsx` - 単一シートのExcelファイル
- `multi_sheet_sample.xlsx` - 複数シートのExcelファイル

## ステップ実行とデバッグ手法

### 1. ブレークポイントの効果的な設置

#### 推奨設置箇所

**データ作成フェーズ:**

```csharp
static List<string[]> CreateSampleData()
{
    var data = new List<string[]>(); // ← ブレークポイント1
    
    string[] headers = { "社員ID", "名前", "年齢", "部署", "給与", "入社日" };
    data.Add(headers); // ← ブレークポイント2
    
    foreach (var employee in employees)
    {
        data.Add(employee); // ← ブレークポイント3（ループ内）
        Console.WriteLine($"従業員データを追加: {employee[1]} ({employee[3]})");
    }
    
    return data; // ← ブレークポイント4
}
```

**Excel操作フェーズ:**
```csharp
try
{
    ExcelHandler.WriteExcel(excelFilePath, sampleData, "社員リスト"); // ← ブレークポイント5
    
    bool fileExists = File.Exists(excelFilePath); // ← ブレークポイント6
    Console.WriteLine($"ファイル存在: {fileExists}");
}
catch (Exception ex)
{
    Console.WriteLine($"エラーが発生しました: {ex.Message}"); // ← ブレークポイント7
}
```

### 2. デバッグ操作の基本

**キー操作:**
| キー | 操作 | 説明 |
|------|------|------|
| `F5` | 続行 | 次のブレークポイントまで実行 |
| `F10` | ステップオーバー | 次の行に移動（メソッド呼び出しは実行するが中に入らない） |
| `F11` | ステップイン | メソッド内に入る |
| `Shift + F11` | ステップアウト | 現在のメソッドから出る |
| `Shift + F5` | 停止 | デバッグを終了 |
| `Ctrl + Shift + F5` | 再開 | デバッグを再開始 |

### 3. 変数監視の活用

#### 変数パネルでの確認
- **ローカル変数**: 現在のスコープ内の変数
- **ウォッチ**: 特定の式を監視
- **コールスタック**: メソッド呼び出しの履歴

#### 有用なウォッチ式の例
```csharp
// データ構造の確認
data.Count
data[0].Length
employee[1]  // 従業員名

// 計算結果の確認
totalSalary / employeeCount
Math.Round(averageSalary, 0)

// ファイル状態の確認
File.Exists(excelFilePath)
new FileInfo(excelFilePath).Length
```

### 4. 条件付きブレークポイント

特定の条件でのみ停止するブレークポイント：

**設定方法:**
1. ブレークポイントを右クリック
2. 「ブレークポイントの編集」を選択
3. 条件を入力

**有用な条件例:**
```csharp
// 特定のループ回数で停止
i == 3

// 特定の従業員で停止
employee[1] == "田中太郎"

// 給与が一定額以上で停止
long.Parse(row[4]) > 500000

// エラー状況で停止
ex.Message.Contains("Access")
```

### 5. デバッグ時の実践的なテクニック

#### 即座実行ウィンドウの活用
デバッグ中に式を評価：
```csharp
// 変数の内容確認
? data.Count
? employee.Where(e => e[3] == "開発部").Count()

// メソッド呼び出し
? Path.GetFileName(excelFilePath)
? DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
```

#### ログ出力による追跡
```csharp
#if DEBUG
Console.WriteLine($"Debug: 現在処理中 - 行:{i}, 従業員:{row[1]}, 給与:{row[4]}");
#endif
```

### 6. パフォーマンス分析

#### 実行時間の測定
```csharp
var stopwatch = System.Diagnostics.Stopwatch.StartNew();

// 処理の実行
ExcelHandler.WriteExcel(excelFilePath, sampleData, "社員リスト");

stopwatch.Stop();
Console.WriteLine($"Excel書き込み時間: {stopwatch.ElapsedMilliseconds}ms");
```

#### メモリ使用量の確認
```csharp
long memoryBefore = GC.GetTotalMemory(false);

// 処理の実行
var data = ExcelHandler.ReadExcel(excelFilePath);

long memoryAfter = GC.GetTotalMemory(false);
Console.WriteLine($"メモリ使用量増加: {(memoryAfter - memoryBefore) / 1024}KB");
```

---

## プロジェクトの拡張

### 1. モジュール化された構造

#### 推奨フォルダ構造
```
ExcelApp/
├── src/
│   ├── Models/           # データモデル
│   │   ├── Employee.cs
│   │   └── Department.cs
│   ├── Services/         # ビジネスロジック
│   │   ├── ExcelService.cs
│   │   ├── DataService.cs
│   │   └── ValidationService.cs
│   ├── Utils/            # ユーティリティ
│   │   ├── Logger.cs
│   │   ├── ConfigManager.cs
│   │   └── FileHelper.cs
│   ├── Config/           # 設定
│   │   └── AppSettings.cs
│   ├── Program.cs
│   └── ExcelHandler.cs
├── tests/                # テストファイル
│   └── TestData.cs
└── docs/                 # ドキュメント
    └── README.md
```

### 2. データモデルの実装例

#### src/Models/Employee.cs
```csharp
using System;

namespace ExcelApp.Models
{
    public class Employee
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public int Age { get; set; }
        public string Department { get; set; }
        public decimal Salary { get; set; }
        public DateTime HireDate { get; set; }
        
        public Employee() { }
        
        public Employee(string[] data)
        {
            if (data.Length >= 6)
            {
                Id = data[0];
                Name = data[1];
                Age = int.TryParse(data[2], out int age) ? age : 0;
                Department = data[3];
                Salary = decimal.TryParse(data[4], out decimal salary) ? salary : 0;
                HireDate = DateTime.TryParse(data[5], out DateTime date) ? date : DateTime.MinValue;
            }
        }
        
        public string[] ToArray()
        {
            return new string[]
            {
                Id,
                Name,
                Age.ToString(),
                Department,
                Salary.ToString(),
                HireDate.ToString("yyyy-MM-dd")
            };
        }
        
        public override string ToString()
        {
            return $"{Name} ({Department}) - {Salary:C}";
        }
    }
}
```

### 3. サービス層の実装例

#### src/Services/ExcelService.cs
```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using ExcelApp.Models;

namespace ExcelApp.Services
{
    public class ExcelService
    {
        private readonly string _filePath;
        
        public ExcelService(string filePath)
        {
            _filePath = filePath;
        }
        
        public List<Employee> LoadEmployees(string worksheetName = "社員リスト")
        {
            var data = ExcelHandler.ReadExcel(_filePath, worksheetName);
            var employees = new List<Employee>();
            
            // ヘッダー行をスキップ
            for (int i = 1; i < data.Count; i++)
            {
                try
                {
                    employees.Add(new Employee(data[i]));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"行 {i + 1} の読み込みでエラー: {ex.Message}");
                }
            }
            
            return employees;
        }
        
        public void SaveEmployees(List<Employee> employees, string worksheetName = "社員リスト")
        {
            var data = new List<string[]>();
            
            // ヘッダー行を追加
            data.Add(new string[] { "社員ID", "名前", "年齢", "部署", "給与", "入社日" });
            
            // データ行を追加
            foreach (var employee in employees)
            {
                data.Add(employee.ToArray());
            }
            
            ExcelHandler.WriteExcel(_filePath, data, worksheetName);
        }
        
        public EmployeeStatistics CalculateStatistics(List<Employee> employees)
        {
            if (!employees.Any()) return new EmployeeStatistics();
            
            return new EmployeeStatistics
            {
                TotalCount = employees.Count,
                AverageSalary = employees.Average(e => (double)e.Salary),
                MaxSalary = employees.Max(e => e.Salary),
                MinSalary = employees.Min(e => e.Salary),
                DepartmentCounts = employees.GroupBy(e => e.Department)
                    .ToDictionary(g => g.Key, g => g.Count())
            };
        }
    }
    
    public class EmployeeStatistics
    {
        public int TotalCount { get; set; }
        public double AverageSalary { get; set; }
        public decimal MaxSalary { get; set; }
        public decimal MinSalary { get; set; }
        public Dictionary<string, int> DepartmentCounts { get; set; } = new Dictionary<string, int>();
    }
}
```

### 4. VSCode設定の最適化

#### .vscode/settings.json
```json
{
    "files.defaultLanguage": "csharp",
    "files.exclude": {
        "**/bin": true,
        "**/packages": true,
        "**/temp": true
    },
    "search.exclude": {
        "**/bin": true,
        "**/packages": true
    },
    "editor.formatOnSave": true,
    "editor.insertSpaces": true,
    "editor.tabSize": 4
}
```

#### コードスニペットの定義
`.vscode/csharp.code-snippets`
```json
{
    "Excel Handler Method": {
        "prefix": "excelmethod",
        "body": [
            "/// <summary>",
            "/// $1",
            "/// </summary>",
            "/// <param name=\"$2\">$3</param>",
            "/// <returns>$4</returns>",
            "public static $5 $6($7)",
            "{",
            "    try",
            "    {",
            "        $8",
            "    }",
            "    catch (Exception ex)",
            "    {",
            "        Console.WriteLine($\"エラー: {ex.Message}\");",
            "        $9",
            "    }",
            "}"
        ],
        "description": "Excel操作メソッドのテンプレート"
    },
    "Employee Model": {
        "prefix": "employee",
        "body": [
            "public class Employee",
            "{",
            "    public string Id { get; set; }",
            "    public string Name { get; set; }",
            "    public int Age { get; set; }",
            "    public string Department { get; set; }",
            "    public decimal Salary { get; set; }",
            "    public DateTime HireDate { get; set; }",
            "",
            "    public Employee() { }",
            "",
            "    public Employee(string[] data)",
            "    {",
            "        // データ配列からプロパティを設定",
            "        $1",
            "    }",
            "}"
        ],
        "description": "従業員モデルクラスのテンプレート"
    }
}
```

---

## トラブルシューティング

### 1. ビルドエラーの対処

#### csc.exeが見つからない場合

**エラーメッセージ:**

```
ERROR: csc.exe not found!
Please ensure .NET Framework is installed.
```

**解決方法:**

1. .NET Frameworkのインストール確認
2. 環境変数PATHの設定確認
3. 手動パス指定

```batch
:: 手動でcsc.exeのパスを確認
dir "C:\Windows\Microsoft.NET\Framework*\v4*\csc.exe" /s
```

#### WindowsBase.dllが見つからない場合

**エラーメッセージ:**
```
error CS0006: Metadata file 'WindowsBase.dll' could not be found
```

**解決方法:**
```batch
:: WindowsBase.dllの場所を確認
dir "C:\Program Files*\Reference Assemblies\Microsoft\Framework\.NETFramework\v4*\WindowsBase.dll" /s

:: 見つからない場合は直接パスを指定
/reference:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\WindowsBase.dll"
```

### 2. PowerShell実行エラーの対処

#### 実行ポリシーエラー
**エラーメッセージ:**
```
execution of scripts is disabled on this system
```

**解決方法:**
```powershell
# 現在のユーザーのみ変更（推奨）
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# 一時的なバイパス
powershell -ExecutionPolicy Bypass -File "setup-libraries.ps1"
```

#### ネットワーク接続エラー
**エラーメッセージ:**
```
Failed to download package: The remote name could not be resolved
```

**解決方法:**
1. インターネット接続の確認
2. プロキシ設定の確認
3. 手動ダウンロードへのフォールバック

```powershell
# プロキシ設定（企業環境）
$proxy = New-Object System.Net.WebProxy("http://proxy.company.com:8080")
$webClient.Proxy = $proxy
```

### 3. デバッグエラーの対処

#### ブレークポイントが無効になる
**症状:** ブレークポイントが灰色で表示される

**原因と解決方法:**
1. **デバッグ情報不足**: `build-debug.bat`を使用してビルド
2. **最適化有効**: `/optimize-`オプションが設定されているか確認
3. **ファイルパス不一致**: ソースファイルの場所を確認

#### 変数値が表示されない
**症状:** 変数にカーソルを合わせても値が表示されない

**解決方法:**
1. デバッグビルドを使用
2. 最適化を無効化
3. 変数のスコープ内であることを確認

### 4. Excel操作エラーの対処

#### ファイルアクセスエラー
**エラーメッセージ:**
```
The process cannot access the file because it is being used by another process
```

**解決方法:**
```csharp
// Excelファイルが開かれていないか確認
// usingステートメントで確実にファイルを閉じる
using (var document = SpreadsheetDocument.Open(filePath, false))
{
    // 処理
} // ここで自動的にファイルが閉じられる
```

#### メモリ不足エラー
**大容量ファイルの処理時:**
```csharp
// ストリーミング処理の実装
public static IEnumerable<string[]> ReadExcelStreaming(string filePath)
{
    using (var document = SpreadsheetDocument.Open(filePath, false))
    {
        var workbookPart = document.WorkbookPart;
        var worksheetPart = workbookPart.WorksheetParts.First();
        var reader = OpenXmlReader.Create(worksheetPart);
        
        while (reader.Read())
        {
            if (reader.ElementType == typeof(Row))
            {
                // 行単位で処理してメモリ使用量を抑制
                yield return ProcessRow(reader.LoadCurrentElement() as Row);
            }
        }
    }
}
```

### 5. 環境固有の問題

#### 企業ファイアウォール環境
**NuGetパッケージダウンロード失敗:**
```powershell
# 企業プロキシ設定
$webClient = New-Object System.Net.WebClient
$webClient.Proxy = [System.Net.WebRequest]::DefaultWebProxy
$webClient.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
```

#### 古いWindowsバージョン
**.NET Framework 4.0以下の環境:**
```batch
:: より古いバージョンのcsc.exeを検索
set "POSSIBLE_PATHS=C:\Windows\Microsoft.NET\Framework64\v3.5\csc.exe"
set "POSSIBLE_PATHS=!POSSIBLE_PATHS! C:\Windows\Microsoft.NET\Framework\v3.5\csc.exe"
```

### 6. ログとデバッグ情報の活用

#### 詳細ログの有効化
```csharp
public class Logger
{
    private static string logPath = "debug.log";
    
    public static void Log(string message)
    {
        var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        var logEntry = $"[{timestamp}] {message}";
        
        Console.WriteLine(logEntry);
        File.AppendAllText(logPath, logEntry + Environment.NewLine);
    }
    
    public static void LogError(Exception ex)
    {
        Log($"ERROR: {ex.Message}");
        Log($"STACK: {ex.StackTrace}");
    }
}
```

#### パフォーマンス測定
```csharp
public class PerformanceTracker
{
    private static Dictionary<string, System.Diagnostics.Stopwatch> timers = 
        new Dictionary<string, System.Diagnostics.Stopwatch>();
    
    public static void StartTimer(string name)
    {
        timers[name] = System.Diagnostics.Stopwatch.StartNew();
    }
    
    public static void StopTimer(string name)
    {
        if (timers.ContainsKey(name))
        {
            timers[name].Stop();
            Logger.Log($"PERF: {name} took {timers[name].ElapsedMilliseconds}ms");
        }
    }
}
```

---

## まとめ

### 実現された機能

このガイドにより、以下の完全な開発環境が構築されました：

#### **技術的成果**
- 🛠️ **ゼロインストール開発環境**: Visual Studio不要のC#開発
- 📊 **本格Excel操作**: DocumentFormat.OpenXmlによる高機能実装
- 🔧 **自動化システム**: NuGetパッケージからビルドまで完全自動化
- 🐛 **フルデバッグ環境**: ブレークポイント、ステップ実行、変数監視
- 🏗️ **プロフェッショナル構造**: 実用的なプロジェクト構成

#### **開発効率の向上**
- ⚡ **ワンクリックセットアップ**: `setup-project.bat`による即座の環境構築
- 🔄 **継続的開発サイクル**: 編集→ビルド→デバッグの高速サイクル
- 📁 **整理された構造**: ソースコード、ライブラリ、出力の明確な分離
- 🔍 **強力なデバッグ**: 条件付きブレークポイント、パフォーマンス測定

#### **企業環境対応**
- 🔐 **セキュリティ遵守**: 管理者権限不要、追加インストール不要
- 🌐 **ネットワーク制約対応**: プロキシ環境での動作保証
- 📋 **標準化**: チーム開発での一貫した環境構築
- 🔧 **カスタマイズ**: 企業固有要件への対応可能

### 適用可能な開発シーン

#### **学習・教育目的**
- C#プログラミングの学習
- Excel操作APIの理解
- デバッグ技術の習得
- .NET Frameworkの理解

#### **業務アプリケーション開発**
- 帳票生成システム
- データ変換ツール
- 定期レポート作成
- 在庫管理システム

#### **プロトタイプ開発**
- 概念実証（PoC）
- アイデア検証
- 顧客デモ用アプリ
- 技術評価ツール

### 今後の拡張可能性

#### **機能拡張**
- データベース連携（SQL Server、SQLite）
- WebAPI連携（REST、SOAP）
- 他のOffice文書対応（Word、PowerPoint）
- クラウドストレージ連携

#### **技術進化への対応**
- .NET 5/6/7への移行準備
- クロスプラットフォーム対応
- コンテナ化対応
- CI/CD パイプライン統合

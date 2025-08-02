# setup-package.ps1
# NuGet パッケージの自動ダウンロードと展開スクリプト

param(
    [string]$PackageName = "",
    [string]$Version = "latest",
    [string]$TargetFramework = "net48"
)

# PackageName の必須チェック
if ([string]::IsNullOrWhiteSpace($PackageName)) {
    Write-Error "ERROR: PackageName is required!"
    Write-Host ""
    Write-Host "Usage: .\setup-packages.ps1 -PackageName <package-name> [-Version <version>] [-TargetFramework <framework>]"
    Write-Host ""
    Write-Host "Examples:"
    Write-Host "  .\setup-packages.ps1 -PackageName 'DocumentFormat.OpenXml'"
    Write-Host "  .\setup-packages.ps1 -PackageName 'Newtonsoft.Json' -Version '13.0.3'"
    Write-Host "  .\setup-packages.ps1 -PackageName 'System.IO.Packaging' -TargetFramework 'net48'"
    exit 1
}

# 設定
$PackagesDir = "packages"
$LibDir = "lib"
$TempDir = "temp"

#Write-Host "=== NuGet Package setup start ==="
Write-Host "Package: $PackageName"
Write-Host "Version: $Version"
Write-Host "Target Framework: $TargetFramework"

# 必要なディレクトリの作成
@($PackagesDir, $LibDir, $TempDir) | ForEach-Object {
    if (!(Test-Path $_)) {
        New-Item -ItemType Directory -Path $_ -Force | Out-Null
        Write-Host "Created directory: $_"
    }
}

# NuGet API からパッケージ情報を取得
function Get-LatestPackageVersion {
    param([string]$PackageName)
    
    try {
        Write-Host "Fetching package information from NuGet API..."
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
        Write-Host "Latest version found: $actualVersion"
        $Version = $actualVersion
    }
} else {
    Write-Host "Using specified version: $Version"
}

# パッケージのダウンロード
$packageFileName = "$PackageName.$Version.nupkg"
$downloadUrl = "https://api.nuget.org/v3-flatcontainer/$($PackageName.ToLower())/$Version/$($PackageName.ToLower()).$Version.nupkg"
$downloadPath = Join-Path $TempDir $packageFileName

Write-Host "Downloading package from: $downloadUrl"

try {
    # WebClient を使用してダウンロード進捗を表示
    $progressActivity = "Downloading $packageFileName"
    
    # プログレスバーの初期表示
    Write-Progress -Activity $progressActivity -Status "Initializing..." -PercentComplete 0
    
    # WebClient でダウンロード実行
    try {
        $webClient = New-Object System.Net.WebClient
        
        # ダウンロード進捗イベントハンドラーを登録
        $progressHandler = {
            param($sender, $e)
            if ($e.TotalBytesToReceive -gt 0) {
                $percentComplete = [math]::Round(($e.BytesReceived / $e.TotalBytesToReceive) * 100, 1)
                $status = "Downloaded $([math]::Round($e.BytesReceived / 1MB, 2)) MB / $([math]::Round($e.TotalBytesToReceive / 1MB, 2)) MB ($percentComplete%)"
                Write-Progress -Activity $progressActivity -Status $status -PercentComplete $percentComplete
            }
        }
        
        # イベントハンドラーを登録
        Register-ObjectEvent -InputObject $webClient -EventName DownloadProgressChanged -Action $progressHandler | Out-Null
        
        # 同期ダウンロード実行
        $webClient.DownloadFile($downloadUrl, $downloadPath)
        
        # イベントハンドラーをクリーンアップ
        Get-EventSubscriber | Where-Object { $_.SourceObject -eq $webClient } | Unregister-Event
        $webClient.Dispose()
        
        Write-Progress -Activity $progressActivity -Status "Download completed" -PercentComplete 100 -Completed
    }
    catch [System.Net.WebException] {
        Write-Progress -Activity $progressActivity -Completed
        if ($_.Exception.Response -and $_.Exception.Response.StatusCode -eq 404) {
            throw "Package '$PackageName' version '$Version' not found on NuGet. Please check the package name and version."
        }
        else {
            throw "Network error occurred: $($_.Exception.Message)"
        }
    }
    catch {
        Write-Progress -Activity $progressActivity -Completed
        throw "Failed to download package: $($_.Exception.Message)"
    }
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

Write-Host "Package downloaded successfully: $downloadPath"
Write-Host "File size: $([math]::Round((Get-Item $downloadPath).Length / 1MB, 2)) MB"

# パッケージの展開
$extractPath = Join-Path $PackagesDir $PackageName

Write-Host "Extracting package to: $extractPath"

try {
    # 既存の展開フォルダを削除
    if (Test-Path $extractPath) {
        Remove-Item $extractPath -Recurse -Force
        Write-Host "Removed existing extraction directory"
    }
    
    # ZIP として展開（.nupkg は ZIP ファイル）
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($downloadPath, $extractPath)
    
    Write-Host "Package extracted successfully"
}
catch {
    Write-Error "Failed to extract package: $($_.Exception.Message)"
    exit 1
}

# DLL ファイルの検索とコピー
Write-Host "Searching for DLL files..."

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
        Write-Host "Checking: $searchPath"
        
        $dllFiles = Get-ChildItem -Path $searchPath -Filter "*.dll" -File
        
        if ($dllFiles.Count -gt 0) {
            Write-Host "Found $($dllFiles.Count) DLL file(s) in $searchPath"
            
            foreach ($dll in $dllFiles) {
                $destPath = Join-Path $LibDir $dll.Name
                Copy-Item $dll.FullName $destPath -Force
                Write-Host "  Copied: $($dll.Name) -> lib\"
                $dllsCopied++
            }
            
            # 主要な DLL が見つかったら他のパスは検索しない
            if ($dllFiles.Name -contains $primaryDll) {
                Write-Host "Primary DLL found, skipping other target frameworks"
                break
            }
        }
    }
}

if ($dllsCopied -eq 0) {
    # Microsoft.Net.Compilers は DLL を含まないため、この警告機能はコメント化
    #Write-Warning "No DLL files were found or copied!"
    #Write-Host "Available lib directories:"
    #
    #$libBasePath = Join-Path $extractPath "lib"
    #if (Test-Path $libBasePath) {
    #    Get-ChildItem $libBasePath -Directory | ForEach-Object {
    #        Write-Host "  - lib\$($_.Name)"
    #    }
    #}
} else {
    Write-Host "Successfully copied $dllsCopied DLL file(s) to lib directory"
}

# 依存関係の確認
$nuspecPath = Join-Path $extractPath "$PackageName.nuspec"
if (Test-Path $nuspecPath) {
    Write-Host "Checking dependencies..."
    
    try {
        [xml]$nuspec = Get-Content $nuspecPath
        $dependencies = $nuspec.package.metadata.dependencies.dependency
        
        if ($dependencies) {
            Write-Host "Dependencies found:"
            $dependencies | ForEach-Object {
                if ($_.id -and $_.version) {
                    Write-Host "  - $($_.id) (>= $($_.version))"
                }
            }
            Write-Host "Note: Dependencies are not automatically downloaded. Install manually if needed."
        } else {
            Write-Host "No dependencies found"
        }
    }
    catch {
        Write-Warning "Failed to parse .nuspec file: $($_.Exception.Message)"
    }
}

# 一時ファイルのクリーンアップ
Write-Host "Cleaning up temporary files..."
Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue

# 結果の表示
Write-Host "=== Setup Summary ==="
Write-Host "Package: $PackageName v$Version"
Write-Host "DLLs copied: $dllsCopied"
Write-Host "Target directory: $LibDir"

#if (Test-Path (Join-Path $LibDir "$PackageName.dll")) {
#    Write-Host "Setup completed successfully!"
#    Write-Host "You can now build your project."
#} else {
#    Write-Warning "Setup may not have completed successfully."
#    Write-Host "Please check the lib directory manually."
#}

#Write-Host "`nPress any key to continue..."
#$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

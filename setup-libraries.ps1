# setup-libraries.ps1
# NuGetパッケージの自動ダウンロードと展開スクリプト

param(
    [string]$PackageName = "DocumentFormat.OpenXml",
    [string]$Version = "latest",
    [string]$TargetFramework = "net481"
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

#Write-Host "`nPress any key to continue..." -ForegroundColor Gray
#$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

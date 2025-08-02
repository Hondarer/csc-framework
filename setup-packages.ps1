# setup-packages.ps1
# packages.config を使った NuGet パッケージの自動ダウンロードと展開スクリプト

$configFile = 'packages.config'

if (Test-Path $configFile) {
    [xml]$config = Get-Content $configFile
    $packages = $config.packages.package
    $totalPackages = $packages.Count
    $currentPackage = 0

    foreach ($package in $packages) {
        $currentPackage++;
        Write-Host "`n[Package $currentPackage/$totalPackages] Processing: $($package.id)"
        $version = if ($package.version -eq 'latest') { 'latest' } else { $package.version }
        $targetFramework = if ($package.targetFramework) { $package.targetFramework } else { 'net48' }
        try {
            & '.\setup-package.ps1' -PackageName $package.id -Version $version -TargetFramework $targetFramework -ErrorAction Stop
            if (-not $?) {
                Write-Error "ERROR: Failed to setup package: $($package.id)"
                exit 1;
            }
        } catch {
            Write-Error "ERROR: Exception occurred while setting up package: $($package.id)"
            Write-Error "Error: $($_.Exception.Message)"
            exit 1
        }
    }
} else {
    Write-Error "ERROR: packages.config not found!"
    exit 1
}

$configFile = 'packages.config';
if (Test-Path $configFile) {
    [xml]$config = Get-Content $configFile;
    $packages = $config.packages.package;
    $totalPackages = $packages.Count;
    $currentPackage = 0;
    foreach ($package in $packages) {
        $currentPackage++;
        Write-Host \"[Package $currentPackage/$totalPackages] Processing: $($package.id)\" -ForegroundColor Yellow;
        $version = if ($package.version -eq 'latest') { 'latest' } else { $package.version };
        $targetFramework = if ($package.targetFramework) { $package.targetFramework } else { 'net481' };
        try {
            & '.\setup-package.ps1' -PackageName $package.id -Version $version -TargetFramework $targetFramework -ErrorAction Stop;
            if (-not $?) {
                Write-Host \"ERROR: Failed to setup package: $($package.id)\" -ForegroundColor Red;
                exit 1;
            }
        } catch {
            Write-Host \"ERROR: Exception occurred while setting up package: $($package.id)\" -ForegroundColor Red;
            Write-Host \"Error: $($_.Exception.Message)\" -ForegroundColor Red;
            exit 1;
        }
    }
} else {
    Write-Host 'ERROR: packages.config not found!' -ForegroundColor Red;
    exit 1;
}

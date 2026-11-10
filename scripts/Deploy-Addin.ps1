$ErrorActionPreference = "Stop"

$projectRoot = "D:\Projects\XLSmart-Development"
$source = Join-Path $projectRoot "workbook\XLSmart_User.xlam"
$templateSource = Join-Path $projectRoot "assets\PowerDashboardTemplate.html"
$appdata = [System.Environment]::GetFolderPath('ApplicationData')
$addinsFolder = Join-Path $appdata "Microsoft\AddIns"
$destination = Join-Path $addinsFolder "XLSmart_User.xlam"
$templateDestination = Join-Path $addinsFolder "PowerDashboardTemplate.html"
$updaterSource = Join-Path $projectRoot "scripts\Install-AddinUpdate.ps1"
$updaterDestination = Join-Path $addinsFolder "Install-AddinUpdate.ps1"

Write-Host "Updating local Add-in in Excel..." -ForegroundColor Cyan
Write-Host "Source: $source"
Write-Host "Destination: $destination"
Write-Host "Template: $templateDestination"

if (-not (Test-Path $source)) {
    Write-Host "[ERROR] Source workbook not found!" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $templateSource)) {
    Write-Host "[ERROR] Dashboard template not found: $templateSource" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path -LiteralPath $updaterSource -PathType Leaf)) {
    Write-Host "[ERROR] Update helper not found: $updaterSource" -ForegroundColor Red
    exit 1
}

# Never replace an add-in while any Excel process may have it loaded.
$excelProc = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
if ($excelProc.Count -gt 0) {
    Write-Host "[ERROR] Excel is running; no files were changed." -ForegroundColor Red
    Write-Host "Close all Excel windows and run this script again." -ForegroundColor Yellow
    exit 2
}

try {
    foreach ($path in @($destination, $templateDestination, $updaterDestination)) {
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            $backup = "$path.backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')_$([Guid]::NewGuid().ToString('N').Substring(0,8))"
            Copy-Item -LiteralPath $path -Destination $backup
        }
    }

    Copy-Item -LiteralPath $source -Destination $destination -Force
    Copy-Item -LiteralPath $templateSource -Destination $templateDestination -Force
    Copy-Item -LiteralPath $updaterSource -Destination $updaterDestination -Force

    $sourceHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
    $deployedHash = (Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash
    if ($sourceHash -ne $deployedHash) {
        throw "Installed add-in hash differs from source. Expected $sourceHash; got $deployedHash."
    }
    Write-Host "[SUCCESS] Add-in successfully updated." -ForegroundColor Green
    Write-Host "[SUCCESS] Dashboard template and updater deployed." -ForegroundColor Green
    Write-Host "New SHA-256: $deployedHash"
} catch {
    Write-Host "[ERROR] Could not deploy files: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Check the timestamped backups in $addinsFolder before retrying." -ForegroundColor Yellow
    exit 1
}
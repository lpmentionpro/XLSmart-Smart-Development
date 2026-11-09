$ErrorActionPreference = "Stop"

$projectRoot = "D:\Projects\XLSmart-Development"
$source = Join-Path $projectRoot "workbook\XLSmart_User.xlam"
$templateSource = Join-Path $projectRoot "assets\PowerDashboardTemplate.html"
$appdata = [System.Environment]::GetFolderPath('ApplicationData')
$addinsFolder = Join-Path $appdata "Microsoft\AddIns"
$destination = Join-Path $addinsFolder "XLSmart_User.xlam"
$templateDestination = Join-Path $addinsFolder "PowerDashboardTemplate.html"

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

# Check if Excel is running
$excelProc = Get-Process excel -ErrorAction SilentlyContinue
if ($excelProc) {
    Write-Host "[WARN] Excel is currently running. Attempting copy..." -ForegroundColor Yellow
}

try {
    Copy-Item -Path $source -Destination $destination -Force
    Copy-Item -Path $templateSource -Destination $templateDestination -Force
    $deployedHash = (Get-FileHash -Path $destination -Algorithm SHA256).Hash
    Write-Host "[SUCCESS] Add-in successfully updated." -ForegroundColor Green
    Write-Host "[SUCCESS] Dashboard template deployed to $templateDestination" -ForegroundColor Green
    Write-Host "New SHA-256: $deployedHash"
} catch {
    Write-Host "[ERROR] Could not copy file: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Please close Excel and run this script again." -ForegroundColor Yellow
}
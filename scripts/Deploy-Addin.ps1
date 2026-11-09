$ErrorActionPreference = "Stop"

$source = "D:\Projects\XLSmart-Development\workbook\XLSmart_User.xlam"
$appdata = [System.Environment]::GetFolderPath('ApplicationData')
$addinsFolder = Join-Path $appdata "Microsoft\AddIns"
$destination = Join-Path $addinsFolder "XLSmart_User.xlam"

Write-Host "Updating local Add-in in Excel..." -ForegroundColor Cyan
Write-Host "Source: $source"
Write-Host "Destination: $destination"

if (-not (Test-Path $source)) {
    Write-Host "[ERROR] Source workbook not found!" -ForegroundColor Red
    exit 1
}

# Check if Excel is running
$excelProc = Get-Process excel -ErrorAction SilentlyContinue
if ($excelProc) {
    Write-Host "[WARN] Excel is currently running. Attempting copy..." -ForegroundColor Yellow
}

try {
    Copy-Item -Path $source -Destination $destination -Force
    $deployedHash = (Get-FileHash -Path $destination -Algorithm SHA256).Hash
    Write-Host "[SUCCESS] Add-in successfully updated." -ForegroundColor Green
    Write-Host "New SHA-256: $deployedHash"
} catch {
    Write-Host "[ERROR] Could not copy file: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Please close Excel and run this script again." -ForegroundColor Yellow
}
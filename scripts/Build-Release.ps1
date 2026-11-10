#Requires -Version 5.1
<#
.SYNOPSIS
    Validates and packages the XLSmart release artifacts.

.DESCRIPTION
    Performs pre-release validation, computes SHA-256 hashes,
    prepares version.json, and copies artifacts to release\.
    Does NOT push to GitHub. Always asks before replacing existing release files.

.PARAMETER Version
    The release version string, e.g. "1.0.3".

.PARAMETER WorkbookPath
    Full path to the built XLSmart_User.xlam.
    Default: workbook\XLSmart_User.xlam

.PARAMETER SourcePath
    Root directory of VBA source files.
    Default: src\

.PARAMETER VersionJsonPath
    Full path to version.json to update.
    Default: installer\version.json

.PARAMETER DownloadBaseUrl
    Base URL for the download_url in version.json.
    Default: https://raw.githubusercontent.com/lpmentionpro/XLSmart-Smart-Development/main/release/

.PARAMETER SkipVbaValidation
    Skip checking that VBA source files are present.

.PARAMETER Force
    Overwrite existing release artifacts without prompting.

.EXAMPLE
    .\Build-Release.ps1 -Version "1.0.3"

.EXAMPLE
    .\Build-Release.ps1 -Version "1.0.3" -Force

.NOTES
    Author: XLSmart Development Team
    Version: 1.0.0
    SAFETY: Does not push to GitHub. Does not modify the original AddIn.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string]$Version,

    [Parameter(Mandatory = $false)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $false)]
    [string]$SourcePath,

    [Parameter(Mandatory = $false)]
    [string]$VersionJsonPath,

    [Parameter(Mandatory = $false)]
    [string]$DownloadBaseUrl = "https://raw.githubusercontent.com/lpmentionpro/XLSmart-Smart-Development/main/release/",

    [switch]$SkipVbaValidation,
    [switch]$Force
)

Set-StrictMode -Off
$ErrorActionPreference = "Stop"

# ─── Helpers ─────────────────────────────────────────────────────────────────

function Write-Banner {
    param([string]$Text)
    $Line = "=" * 70
    Write-Host "`n$Line" -ForegroundColor Cyan
    Write-Host " $Text" -ForegroundColor Cyan
    Write-Host "$Line" -ForegroundColor Cyan
}

function Write-Pass { param([string]$M) Write-Host "  [PASS] $M" -ForegroundColor Green }
function Write-Fail { param([string]$M) Write-Host "  [FAIL] $M" -ForegroundColor Red; $script:FailCount++ }
function Write-Warn { param([string]$M) Write-Host "  [WARN] $M" -ForegroundColor Yellow }
function Write-Info { param([string]$M) Write-Host "  [ >> ] $M" -ForegroundColor White }

$script:FailCount = 0

function Require-FileExists {
    param([string]$Path, [string]$Label)
    if (Test-Path $Path) { Write-Pass "$Label found: $Path" }
    else { Write-Fail "$Label NOT FOUND: $Path" }
}

function Get-SHA256 {
    param([string]$Path)
    return (Get-FileHash $Path -Algorithm SHA256).Hash
}

function Get-FileSize {
    param([string]$Path)
    $Bytes = (Get-Item $Path).Length
    return "{0:N0} bytes ({1:N1} KB)" -f $Bytes, ($Bytes / 1024)
}

# ─── Path resolution ─────────────────────────────────────────────────────────

$ScriptDir   = Split-Path $MyInvocation.MyCommand.Path -Parent
$ProjectRoot = Split-Path $ScriptDir -Parent

if ([string]::IsNullOrEmpty($WorkbookPath)) {
    $WorkbookPath = Join-Path $ProjectRoot "workbook\XLSmart_User.xlam"
}

if ([string]::IsNullOrEmpty($SourcePath)) {
    $SourcePath = Join-Path $ProjectRoot "src"
}

if ([string]::IsNullOrEmpty($VersionJsonPath)) {
    $VersionJsonPath = Join-Path $ProjectRoot "installer\version.json"
}

$ReleasePath = Join-Path $ProjectRoot "release"
$BackupPath  = Join-Path $ProjectRoot "backups"

if ($DownloadBaseUrl -notmatch '/$') { $DownloadBaseUrl += "/" }

# ─── Step 1: Validate source files ───────────────────────────────────────────

Write-Banner "Step 1: Source File Validation"

$RequiredModules = @(
    "modPowerDashboardEngine",
    "modRibbon",
    "modRibbonCallbacks",
    "modSmartDataManager",
    "modVersion"
)

if (-not $SkipVbaValidation) {
    $ModulesPath = Join-Path $SourcePath "Modules"
    $ClassesPath = Join-Path $SourcePath "Classes"
    $FormsPath   = Join-Path $SourcePath "Forms"

    $ModuleCount  = (Get-ChildItem $ModulesPath -Filter "*.bas" -ErrorAction SilentlyContinue).Count
    $ClassCount   = (Get-ChildItem $ClassesPath -Filter "*.cls" -ErrorAction SilentlyContinue).Count
    $FormCount    = (Get-ChildItem $FormsPath   -Filter "*.frm" -ErrorAction SilentlyContinue).Count
    $FrxCount     = (Get-ChildItem $FormsPath   -Filter "*.frx" -ErrorAction SilentlyContinue).Count

    if ($ModuleCount -ge 15)  { Write-Pass "Modules:   $ModuleCount .bas files" }
    else                      { Write-Fail "Modules:   only $ModuleCount .bas files (expected ≥ 15)" }

    if ($ClassCount  -ge 2)   { Write-Pass "Classes:   $ClassCount .cls files" }
    else                      { Write-Fail "Classes:   only $ClassCount .cls files (expected ≥ 2)" }

    if ($FormCount   -ge 8)   { Write-Pass "Forms:     $FormCount .frm files" }
    else                      { Write-Fail "Forms:     only $FormCount .frm files (expected ≥ 8)" }

    if ($FrxCount -eq $FormCount) { Write-Pass "FRX:       $FrxCount .frx files (matches .frm count)" }
    else                          { Write-Fail "FRX:       $FrxCount .frx files (expected $FormCount to match .frm count)" }

    foreach ($Mod in $RequiredModules) {
        Require-FileExists (Join-Path $ModulesPath "$Mod.bas") $Mod
    }
}
else {
    Write-Warn "VBA source validation skipped (-SkipVbaValidation)"
}

# ─── Step 2: Validate workbook ───────────────────────────────────────────────

Write-Banner "Step 2: Workbook Validation"

if (-not (Test-Path $WorkbookPath)) {
    Write-Fail "Workbook not found: $WorkbookPath"
}
else {
    $WorkbookSize = Get-FileSize $WorkbookPath
    $WorkbookHash = Get-SHA256 $WorkbookPath

    Write-Pass "Workbook:  $WorkbookPath"
    Write-Info "Size:      $WorkbookSize"
    Write-Info "SHA-256:   $WorkbookHash"

    # Verify it's not the AddIn directory copy
    $AddInPath = [System.Environment]::ExpandEnvironmentVariables("%APPDATA%\Microsoft\AddIns")
    if ($WorkbookPath.ToLower().StartsWith($AddInPath.ToLower())) {
        Write-Fail "Workbook is inside AddIns folder — targeting original add-in is forbidden."
    }
    else {
        Write-Pass "Workbook is NOT in AddIns — safe to package."
    }
}

# ─── Step 3: Stop if validation failed ───────────────────────────────────────

if ($script:FailCount -gt 0) {
    Write-Host "`n[ABORT] $($script:FailCount) validation failure(s) detected. Fix these before building a release." -ForegroundColor Red
    exit 1
}

# ─── Step 4: Prepare release folder ──────────────────────────────────────────

Write-Banner "Step 3: Prepare Release Artifacts"

if (-not (Test-Path $ReleasePath)) {
    New-Item -ItemType Directory -Path $ReleasePath -Force | Out-Null
    Write-Info "Created release\"
}

$ReleaseXlam = Join-Path $ReleasePath "XLSmart_User.xlam"
$UpdaterSource = Join-Path $ScriptDir "Install-AddinUpdate.ps1"
$UpdaterDestination = Join-Path $ReleasePath "Install-AddinUpdate.ps1"

if (-not (Test-Path -LiteralPath $UpdaterSource -PathType Leaf)) {
    Write-Fail "Updater helper not found: $UpdaterSource"
}

if ((Test-Path $ReleaseXlam) -and -not $Force) {
    Write-Host "`n  Existing release file: $ReleaseXlam" -ForegroundColor Yellow
    $Confirm = Read-Host "  Overwrite? (y/N)"
    if ($Confirm.ToLower() -ne 'y') {
        Write-Host "`n[ABORT] Release cancelled by user." -ForegroundColor Red
        exit 0
    }
}

# Copy workbook and trusted updater helper to release
Write-Info "Copying workbook and updater helper to release\"
Copy-Item $WorkbookPath $ReleaseXlam -Force
Copy-Item -LiteralPath $UpdaterSource -Destination $UpdaterDestination -Force
Write-Pass "Copied: $ReleaseXlam"
Write-Pass "Copied updater helper: $UpdaterDestination"

# Compute hash of release copy (authoritative)
$ReleaseHash = Get-SHA256 $ReleaseXlam
Write-Info "SHA-256 (release copy): $ReleaseHash"

# ─── Step 5: Update version.json ─────────────────────────────────────────────

Write-Banner "Step 4: Update version.json"

$DownloadUrl = "$DownloadBaseUrl/XLSmart_User.xlam"

$VersionObj = [ordered]@{
    version      = $Version
    sha256       = $ReleaseHash.ToLower()
    download_url = $DownloadUrl
    released_at  = (Get-Date -Format "yyyy-MM-dd")
}

$VersionJson = $VersionObj | ConvertTo-Json -Depth 5
$VersionJson | Set-Content $VersionJsonPath -Encoding UTF8

Write-Pass "version.json updated:"
Write-Host ($VersionJson) -ForegroundColor DarkGray

# Also copy version.json to release\
$ReleaseVersionJson = Join-Path $ReleasePath "version.json"
Copy-Item $VersionJsonPath $ReleaseVersionJson -Force
Write-Pass "Copied version.json → release\"

# ─── Step 6: Update CHANGELOG.md header ──────────────────────────────────────

Write-Banner "Step 5: CHANGELOG.md"

$ChangelogPath = Join-Path $ProjectRoot "CHANGELOG.md"
$Today = Get-Date -Format "yyyy-MM-dd"

if (Test-Path $ChangelogPath) {
    $CLContent = Get-Content $ChangelogPath -Raw

    # Move [Unreleased] content under the new version header
    if ($CLContent -match '## \[Unreleased\]') {
        $NewBlock = "## [$Version] - $Today"
        $CLContent = $CLContent -replace '## \[Unreleased\]', "## [Unreleased]`n`n$NewBlock"
        $CLContent | Set-Content $ChangelogPath -Encoding UTF8
        Write-Pass "CHANGELOG.md: Added header [$Version] - $Today"
    }
    else {
        Write-Warn "Could not find [Unreleased] section in CHANGELOG.md — edit manually."
    }

    # Append the release hash only once beneath the version header.
    $HashLine = '### Release Artifacts'
    $HashEntry = '- XLSmart_User.xlam SHA-256: ' + $ReleaseHash.ToLower()
    $CLContent = Get-Content -LiteralPath $ChangelogPath -Raw
    $EscapedVersion = [regex]::Escape($Version)
    $HeaderPattern = '## \[' + $EscapedVersion + '\] - ' + $Today
    if ($CLContent -notmatch [regex]::Escape($HashEntry)) {
        $HashBlock = $HashLine + [Environment]::NewLine + $HashEntry
        $CLContent = $CLContent -replace $HeaderPattern, ('$0' + [Environment]::NewLine + [Environment]::NewLine + $HashBlock)
        $CLContent | Set-Content -LiteralPath $ChangelogPath -Encoding UTF8
    }
    Write-Pass "CHANGELOG.md: Ensured SHA-256 entry for v$Version"
}
else {
    Write-Warn "CHANGELOG.md not found — skipping."
}

# ─── Final summary ────────────────────────────────────────────────────────────

Write-Banner "Build Complete"
Write-Pass  "Version:        $Version"
Write-Info  "Workbook:       $ReleaseXlam"
Write-Info  "SHA-256:        $ReleaseHash"
Write-Info  "Download URL:   $DownloadUrl"
Write-Info  "version.json:   $VersionJsonPath"

Write-Host "`n  ──────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  Release files are in: release\" -ForegroundColor White
Write-Host "  Review changes and SHA-256 in CHANGELOG.md before pushing." -ForegroundColor Yellow
Write-Host "  NEVER push to GitHub without explicit approval." -ForegroundColor Red
Write-Host "  ──────────────────────────────────────────────────────────────`n" -ForegroundColor DarkGray

return [PSCustomObject]@{
    Version      = $Version
    WorkbookPath = $ReleaseXlam
    Hash         = $ReleaseHash
    VersionJson  = $VersionJsonPath
    DownloadUrl  = $DownloadUrl
}

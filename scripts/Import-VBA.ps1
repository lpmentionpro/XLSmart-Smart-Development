#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateScript({
        if (-not (Test-Path $_)) { throw "Workbook not found: $_" }
        if ($_ -notmatch '\.(xlam|xlsm)$') { throw "Must be .xlam or .xlsm" }
        $true
    })]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $false)]
    [string]$SourcePath,

    [Parameter(Mandatory = $false)]
    [string]$BackupPath,

    [Parameter(Mandatory = $false)]
    [string[]]$ComponentFilter,

    [switch]$ForceAll,
    [switch]$NoBackup,
    [switch]$NoCompile,
    [switch]$Quiet
)

$ErrorActionPreference = "Stop"

# Path resolution
$WorkbookPath = (Resolve-Path $WorkbookPath).Path
$ProjectRoot = Split-Path (Split-Path $WorkbookPath -Parent) -Parent

if ([string]::IsNullOrEmpty($SourcePath)) {
    $SourcePath = Join-Path $ProjectRoot "src"
}

if ([string]::IsNullOrEmpty($BackupPath)) {
    $BackupPath = Join-Path $ProjectRoot "backups"
}

$ModulesPath = Join-Path $SourcePath "Modules"
$ClassesPath = Join-Path $SourcePath "Classes"
$FormsPath = Join-Path $SourcePath "Forms"
$DocumentsPath = Join-Path $SourcePath "Documents"

# Safety check - never modify original AddIn
$AddInPath = [System.Environment]::ExpandEnvironmentVariables("%APPDATA%\Microsoft\AddIns")
if ($WorkbookPath.ToLower().StartsWith($AddInPath.ToLower())) {
    Write-Host "[ABORT] Cannot modify original AddIn" -ForegroundColor Red
    exit 1
}

# Statistics
$Stats = @{
    Imported = 0
    Skipped = 0
    Changed = 0
    Errors = @()
    ImportedNames = @()
}

# Build source file map
$SourceMap = @{}

function Add-SourceFiles {
    param([string]$Dir, [string]$Ext, [string]$TypeHint)
    if (-not (Test-Path $Dir)) { return }
    Get-ChildItem $Dir -Filter "*$Ext" | ForEach-Object {
        $Name = $_.BaseName
        $SourceMap[$Name] = @{ Path = $_.FullName; Ext = $Ext; Type = $TypeHint }
    }
}

Add-SourceFiles $ModulesPath ".bas" "Module"
Add-SourceFiles $ClassesPath ".cls" "Class"
Add-SourceFiles $FormsPath ".frm" "Form"
Add-SourceFiles $DocumentsPath ".cls" "Document"

if (-not $Quiet) {
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " XLSmart Import-VBA" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Workbook: $WorkbookPath"
    Write-Host "  Components in source: $($SourceMap.Count)"
}

# Apply filter
if ($ComponentFilter -and $ComponentFilter.Count -gt 0) {
    $Filtered = @{}
    # Support both PowerShell array arguments and comma-separated values passed by a shell.
    $RequestedComponents = $ComponentFilter -split ','
    foreach ($Name in $RequestedComponents) {
        $Name = $Name.Trim()
        if ($SourceMap.ContainsKey($Name)) {
            $Filtered[$Name] = $SourceMap[$Name]
        }
    }
    $SourceMap = $Filtered
    if (-not $Quiet) { Write-Host "  Filtered to: $($SourceMap.Count)" -ForegroundColor Yellow }
}

# Backup workbook
if (-not $NoBackup) {
    if (-not $Quiet) { Write-Host "`n[1] Backing up workbook..." -ForegroundColor Cyan }
    $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $BaseName = [System.IO.Path]::GetFileNameWithoutExtension($WorkbookPath)
    $Ext = [System.IO.Path]::GetExtension($WorkbookPath)

    if (-not (Test-Path $BackupPath)) {
        New-Item -ItemType Directory -Path $BackupPath -Force | Out-Null
    }

    $BackupFile = Join-Path $BackupPath "${BaseName}_${Timestamp}${Ext}"
    Copy-Item $WorkbookPath $BackupFile -Force
    if (-not $Quiet) { Write-Host "  [OK] Backup: $BackupFile" -ForegroundColor Green }
}

# Excel COM automation
$Excel = $null
$Workbook = $null

try {
    if (-not $Quiet) { Write-Host "`n[2] Opening Excel..." -ForegroundColor Cyan }

    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Excel.EnableEvents = $false

    $Workbook = $Excel.Workbooks.Open($WorkbookPath, $false, $false)

    $VBProject = $Workbook.VBProject
    if (-not $VBProject) {
        throw "Cannot access VBProject. Enable Trust access to VBA project object model."
    }

    if (-not $Quiet) { Write-Host "  [OK] VBProject: $($VBProject.Name)" -ForegroundColor Green }

    # Build lookup
    $ExistingComponents = @{}
    foreach ($Comp in $VBProject.VBComponents) {
        $ExistingComponents[$Comp.Name] = $Comp
    }

    if (-not $Quiet) { Write-Host "`n[3] Importing components..." -ForegroundColor Cyan }

    # Import loop
    foreach ($Entry in $SourceMap.GetEnumerator()) {
        $Name = $Entry.Key
        $Info = $Entry.Value
        $SourceFile = $Info.Path
        $TypeHint = $Info.Type

        # Detect document modules
        $IsDocument = ($TypeHint -eq "Document") -or
                      ($Name -eq "ThisWorkbook") -or
                      ($Name -match "^Sheet\d*$")

        # Check if component exists
        if (-not $ExistingComponents.ContainsKey($Name)) {
            # New component - import it
            try {
                $VBProject.VBComponents.Import($SourceFile) | Out-Null
                if (-not $Quiet) { Write-Host "  [NEW] $Name" -ForegroundColor Green }
                $Stats.Imported++
                $Stats.ImportedNames += $Name
            }
            catch {
                $Stats.Errors += "Import failed for $Name : $($_.Exception.Message)"
                if (-not $Quiet) { Write-Host "  [ERR] $Name : $($_.Exception.Message)" -ForegroundColor Red }
            }
            continue
        }

        # Component exists - check if document module
        if ($IsDocument) {
            # Document modules cannot be removed - update code only
            try {
                $Comp = $ExistingComponents[$Name]
                $CodeModule = $Comp.CodeModule

                $Lines = Get-Content $SourceFile
                $CodeLines = @()
                $InCode = $false

                foreach ($Line in $Lines) {
                    if ($Line -match '^Attribute\s+VB_') { continue }
                    if ($Line -match '^VERSION\s+') { continue }
                    if ($Line -match '^\s*$' -and -not $InCode) { continue }
                    $InCode = $true
                    $CodeLines += $Line
                }

                $CodeText = $CodeLines -join "`r`n"

                if ($CodeModule.CountOfLines -gt 0) {
                    $CodeModule.DeleteLines(1, $CodeModule.CountOfLines)
                }

                if ($CodeText.Trim().Length -gt 0) {
                    $CodeModule.InsertLines(1, $CodeText)
                }

                if (-not $Quiet) { Write-Host "  [UPD] $Name (document)" -ForegroundColor Green }
                $Stats.Imported++
                $Stats.ImportedNames += $Name
            }
            catch {
                $Stats.Errors += "Update failed for document $Name : $($_.Exception.Message)"
                if (-not $Quiet) { Write-Host "  [ERR] $Name : $($_.Exception.Message)" -ForegroundColor Red }
            }
            continue
        }

        # Regular module - remove and re-import
        try {
            $ExistingComp = $ExistingComponents[$Name]
            $VBProject.VBComponents.Remove($ExistingComp)
            $VBProject.VBComponents.Import($SourceFile) | Out-Null

            if (-not $Quiet) { Write-Host "  [IMP] $Name" -ForegroundColor Green }
            $Stats.Imported++
            $Stats.ImportedNames += $Name
        }
        catch {
            $Stats.Errors += "Import failed for $Name : $($_.Exception.Message)"
            if (-not $Quiet) { Write-Host "  [ERR] $Name : $($_.Exception.Message)" -ForegroundColor Red }
        }
    }

    # Save
    if ($Stats.Imported -gt 0) {
        if (-not $Quiet) { Write-Host "`n[4] Saving workbook..." -ForegroundColor Cyan }
        $Workbook.Save()
        if (-not $Quiet) { Write-Host "  [OK] Saved" -ForegroundColor Green }
    }

    # Summary
    if (-not $Quiet) {
        Write-Host "`n========================================" -ForegroundColor Cyan
        Write-Host " Import Summary" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "  Imported: $($Stats.Imported)" -ForegroundColor Green
        Write-Host "  Errors:   $($Stats.Errors.Count)" -ForegroundColor $(if ($Stats.Errors.Count -gt 0) { "Red" } else { "Green" })

        if ($Stats.Errors.Count -gt 0) {
            Write-Host "`n  Errors:" -ForegroundColor Red
            $Stats.Errors | ForEach-Object { Write-Host "    - $_" -ForegroundColor Red }
        }
        Write-Host "========================================`n" -ForegroundColor Cyan
    }
}
catch {
    Write-Host "`n[CRITICAL] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    exit 1
}
finally {
    if ($Workbook) {
        try { $Workbook.Close($false) } catch {}
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null
    }
    if ($Excel) {
        try { $Excel.Quit() } catch {}
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

return [PSCustomObject]$Stats

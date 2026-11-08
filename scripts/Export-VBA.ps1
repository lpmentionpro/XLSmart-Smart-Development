#Requires -Version 5.1
<#
.SYNOPSIS
    Exports VBA components from an Excel workbook to source files.

.DESCRIPTION
    Exports all VBA components (standard modules, class modules, user forms,
    document modules) from an Excel .xlam or .xlsm workbook to organized
    source directories. Uses Excel COM automation.

.PARAMETER WorkbookPath
    Full path to the Excel workbook (.xlam or .xlsm).

.PARAMETER OutputPath
    Root directory for exported source files.
    Default: src\ relative to workbook location.

.PARAMETER BackupPath
    Optional backup directory for existing source files.
    Default: backups\export_<timestamp>

.PARAMETER NoBackup
    Skip creating backup of existing source files.

.PARAMETER Quiet
    Suppress progress messages.

.EXAMPLE
    .\Export-VBA.ps1 -WorkbookPath "D:\Projects\XLSmart\workbook\XLSmart_User.xlam"

.EXAMPLE
    .\Export-VBA.ps1 -WorkbookPath ".\workbook\XLSmart_User.xlam" -OutputPath ".\src" -NoBackup

.NOTES
    Author: XLSmart Development Team
    Requires: Excel installed, Trust access to VBA project object model enabled.
    Version: 1.0.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateScript({
        if (-not (Test-Path $_)) {
            throw "Workbook not found: $_"
        }
        if ($_ -notmatch '\.(xlam|xlsm)$') {
            throw "Workbook must be .xlam or .xlsm file"
        }
        $true
    })]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string]$BackupPath,

    [switch]$NoBackup,

    [switch]$Quiet
)

# Resolve paths
$WorkbookPath = (Resolve-Path $WorkbookPath).Path
$ProjectRoot = Split-Path $WorkbookPath -Parent | Split-Path -Parent

if ([string]::IsNullOrEmpty($OutputPath)) {
    $OutputPath = Join-Path $ProjectRoot "src"
}

# Ensure output directories exist
$ModulesPath = Join-Path $OutputPath "Modules"
$ClassesPath = Join-Path $OutputPath "Classes"
$FormsPath = Join-Path $OutputPath "Forms"
$DocumentsPath = Join-Path $OutputPath "Documents"

@($ModulesPath, $ClassesPath, $FormsPath, $DocumentsPath) | ForEach-Object {
    if (-not (Test-Path $_)) {
        New-Item -ItemType Directory -Path $_ -Force | Out-Null
        if (-not $Quiet) { Write-Host "Created directory: $_" -ForegroundColor DarkGray }
    }
}

# Statistics
$Stats = @{
    Modules    = 0
    Classes    = 0
    Forms      = 0
    Documents  = 0
    Total      = 0
    Skipped    = 0
    Errors     = @()
}

# Excel objects
$Excel = $null
$Workbook = $null

try {
    if (-not $Quiet) { Write-Host "`nOpening Excel..." -ForegroundColor Cyan }

    # Create Excel instance
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false

    # Open workbook
    if (-not $Quiet) { Write-Host "Opening workbook: $WorkbookPath" -ForegroundColor Cyan }
    $Workbook = $Excel.Workbooks.Open($WorkbookPath, $false, $true)  # ReadOnly

    # Access VBProject
    $VBProject = $Workbook.VBProject

    if (-not $VBProject) {
        throw "Cannot access VBProject. Ensure 'Trust access to the VBA project object model' is enabled in Excel Trust Center."
    }

    if (-not $Quiet) { Write-Host "VBProject: $($VBProject.Name)" -ForegroundColor Green }

    # Export each component
    foreach ($Component in $VBProject.VBComponents) {
        $ComponentName = $Component.Name
        $ComponentType = $Component.Type

        # Skip empty components
        $CodeModule = $Component.CodeModule
        $LineCount = $CodeModule.CountOfLines

        if ($LineCount -eq 0 -and $ComponentType -ne 100) {
            if (-not $Quiet) { Write-Host "  Skipped (empty): $ComponentName" -ForegroundColor DarkGray }
            $Stats.Skipped++
            continue
        }

        $ExportPath = $null
        $Extension = $null
        $TypeLabel = $null

        switch ($ComponentType) {
            1 {  # vbext_ct_StdModule
                $ExportPath = $ModulesPath
                $Extension = ".bas"
                $TypeLabel = "Module"
                $Stats.Modules++
            }
            2 {  # vbext_ct_ClassModule
                $ExportPath = $ClassesPath
                $Extension = ".cls"
                $TypeLabel = "Class"
                $Stats.Classes++
            }
            3 {  # vbext_ct_MSForm
                $ExportPath = $FormsPath
                $Extension = ".frm"
                $TypeLabel = "Form"
                $Stats.Forms++
            }
            100 {  # vbext_ct_Document
                $ExportPath = $DocumentsPath
                $Extension = ".cls"
                $TypeLabel = "Document"
                $Stats.Documents++
            }
            default {
                if (-not $Quiet) { Write-Host "  Skipped (unknown type): $ComponentName (Type: $ComponentType)" -ForegroundColor Yellow }
                $Stats.Skipped++
                continue
            }
        }

        $TargetFile = Join-Path $ExportPath "$ComponentName$Extension"

        try {
            # Backup existing file if needed
            if ((Test-Path $TargetFile) -and -not $NoBackup) {
                if ([string]::IsNullOrEmpty($BackupPath)) {
                    $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                    $BackupPath = Join-Path $ProjectRoot "backups\export_$Timestamp"
                }

                if (-not (Test-Path $BackupPath)) {
                    New-Item -ItemType Directory -Path $BackupPath -Force | Out-Null
                }

                $BackupDir = Join-Path $BackupPath (Split-Path $ExportPath -Leaf)
                if (-not (Test-Path $BackupDir)) {
                    New-Item -ItemType Directory -Path $BackupDir -Force | Out-Null
                }

                Copy-Item $TargetFile (Join-Path $BackupDir "$ComponentName$Extension") -Force
            }

            # Export component
            $Component.Export($TargetFile)
            $Stats.Total++

            if (-not $Quiet) {
                Write-Host "  Exported [$TypeLabel]: $ComponentName -> $TargetFile" -ForegroundColor Green
            }
        }
        catch {
            $ErrorMsg = "Failed to export $ComponentName : $($_.Exception.Message)"
            $Stats.Errors += $ErrorMsg
            Write-Host "  ERROR: $ErrorMsg" -ForegroundColor Red
        }
    }

    # Summary
    if (-not $Quiet) {
        Write-Host "`n========================================" -ForegroundColor Cyan
        Write-Host "Export Summary" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "  Modules:   $($Stats.Modules)" -ForegroundColor White
        Write-Host "  Classes:   $($Stats.Classes)" -ForegroundColor White
        Write-Host "  Forms:     $($Stats.Forms)" -ForegroundColor White
        Write-Host "  Documents: $($Stats.Documents)" -ForegroundColor White
        Write-Host "  Skipped:   $($Stats.Skipped)" -ForegroundColor DarkGray
        Write-Host "  Total:     $($Stats.Total)" -ForegroundColor Green

        if ($Stats.Errors.Count -gt 0) {
            Write-Host "`nErrors:" -ForegroundColor Red
            $Stats.Errors | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
        }
        Write-Host "========================================`n" -ForegroundColor Cyan
    }
}
catch {
    Write-Host "`nCRITICAL ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    exit 1
}
finally {
    # Cleanup COM objects
    if ($Workbook) {
        $Workbook.Close($false)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null
    }

    if ($Excel) {
        $Excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    }

    # Force garbage collection
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# Return stats as object for programmatic use
return [PSCustomObject]$Stats

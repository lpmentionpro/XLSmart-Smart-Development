#Requires -Version 5.1
<#
.SYNOPSIS
    Compares VBA source files against the development workbook.

.DESCRIPTION
    Exports each component from the workbook to a temp location and
    compares against src\. Reports which components are newer in source,
    newer in workbook, new (only in source), or missing (only in workbook).
    Does NOT modify any files.

.PARAMETER WorkbookPath
    Full path to the Excel workbook (.xlam or .xlsm).

.PARAMETER SourcePath
    Root directory of VBA source files.
    Default: src\ relative to project root.

.PARAMETER OutputReport
    Optional path to write the report as plain text.

.PARAMETER Quiet
    Suppress progress messages; only output final report table.

.EXAMPLE
    .\Compare-VBA.ps1 -WorkbookPath ".\workbook\XLSmart_User.xlam"

.EXAMPLE
    .\Compare-VBA.ps1 -WorkbookPath ".\workbook\XLSmart_User.xlam" -OutputReport ".\docs\compare_report.txt"

.NOTES
    Author: XLSmart Development Team
    Requires: Excel installed, Trust access to VBA project object model enabled.
    Version: 1.0.0
#>

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
    [string]$OutputReport,

    [switch]$Quiet
)

Set-StrictMode -Off
$ErrorActionPreference = "Continue"

# ─── Helpers ─────────────────────────────────────────────────────────────────

function Write-Step {
    param([string]$M, [string]$C = "White")
    if (-not $Quiet) { Write-Host $M -ForegroundColor $C }
}

function Normalize-Lines {
    param([string[]]$Lines)
    return ($Lines | Where-Object {
        $_ -notmatch '^Attribute VB_(Creation|Exposed)ID\s*=' -and
        $_.Trim() -ne ""
    }) -join "`n"
}

# ─── Path resolution ─────────────────────────────────────────────────────────

$WorkbookPath = (Resolve-Path $WorkbookPath).Path
$ProjectRoot  = Split-Path (Split-Path $WorkbookPath -Parent) -Parent

if ([string]::IsNullOrEmpty($SourcePath)) {
    $SourcePath = Join-Path $ProjectRoot "src"
}

$ModulesPath   = Join-Path $SourcePath "Modules"
$ClassesPath   = Join-Path $SourcePath "Classes"
$FormsPath     = Join-Path $SourcePath "Forms"
$DocumentsPath = Join-Path $SourcePath "Documents"

$TempExportDir = Join-Path $env:TEMP "XLSmart_Compare_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -ItemType Directory -Path $TempExportDir -Force | Out-Null

# ─── Collect source files ─────────────────────────────────────────────────────

$SourceMap = @{}
function Register-SourceDir {
    param([string]$Dir, [string]$Ext)
    if (-not (Test-Path $Dir)) { return }
    Get-ChildItem $Dir -Filter "*$Ext" | ForEach-Object {
        $SourceMap[$_.BaseName] = @{ SrcPath = $_.FullName; Ext = $Ext }
    }
}
Register-SourceDir $ModulesPath   ".bas"
Register-SourceDir $ClassesPath   ".cls"
Register-SourceDir $FormsPath     ".frm"
Register-SourceDir $DocumentsPath ".cls"

$Results = [System.Collections.Generic.List[PSCustomObject]]::new()

# ─── Excel COM ───────────────────────────────────────────────────────────────

$Excel    = $null
$Workbook = $null

try {
    Write-Step "`nOpening Excel (read-only)..." "Cyan"
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false

    $Workbook = $Excel.Workbooks.Open($WorkbookPath, $false, $true)
    $VBProject = $Workbook.VBProject

    if (-not $VBProject) {
        throw "Cannot access VBProject. Enable 'Trust access to the VBA project object model'."
    }

    Write-Step "VBProject: $($VBProject.Name)" "Green"
    Write-Step "Exporting workbook components for comparison..." "Cyan"

    $WorkbookNames = [System.Collections.Generic.HashSet[string]]::new()

    # ── Export each workbook component to temp ────────────────────────────────
    foreach ($Comp in $VBProject.VBComponents) {
        $Name = $Comp.Name
        $Type = $Comp.Type
        $null = $WorkbookNames.Add($Name)

        # Determine extension
        $Ext = switch ($Type) {
            1 { ".bas" }   # StdModule
            2 { ".cls" }   # ClassModule
            3 { ".frm" }   # MSForm
            100 { ".cls" } # Document
            default { $null }
        }

        $TypeLabel = switch ($Type) {
            1 { "Module"   }
            2 { "Class"    }
            3 { "Form"     }
            100 { "Document" }
            default { "Unknown" }
        }

        if (-not $Ext) {
            $Results.Add([PSCustomObject]@{
                Name = $Name; Type = "?"; Status = "UNKNOWN_TYPE"; SrcLines = 0; WbkLines = 0; Detail = "Type $Type not handled"
            })
            continue
        }

        $TempFile = Join-Path $TempExportDir "$Name$Ext"

        try {
            $Comp.Export($TempFile)
        }
        catch {
            $Results.Add([PSCustomObject]@{
                Name = $Name; Type = $TypeLabel; Status = "EXPORT_ERROR"; SrcLines = 0; WbkLines = 0
                Detail = $_.Exception.Message
            })
            continue
        }

        # Lines in workbook component
        $WbkLines = (Get-Content $TempFile | Measure-Object -Line).Lines

        # Is there a source file?
        if (-not $SourceMap.ContainsKey($Name)) {
            $Results.Add([PSCustomObject]@{
                Name = $Name; Type = $TypeLabel; Status = "ONLY_IN_WBK"; SrcLines = 0; WbkLines = $WbkLines
                Detail = "Exists in workbook, not in src\"
            })
            continue
        }

        $SrcFile = $SourceMap[$Name].SrcPath
        $SrcLines = (Get-Content $SrcFile | Measure-Object -Line).Lines

        # Normalize and compare
        $SrcNorm = Normalize-Lines (Get-Content $SrcFile)
        $WbkNorm = Normalize-Lines (Get-Content $TempFile)

        if ($SrcNorm -eq $WbkNorm) {
            $Results.Add([PSCustomObject]@{
                Name = $Name; Type = $TypeLabel; Status = "IDENTICAL"; SrcLines = $SrcLines; WbkLines = $WbkLines
                Detail = "No differences"
            })
        }
        else {
            # Count differing lines
            $SrcArr = $SrcNorm -split "`n"
            $WbkArr = $WbkNorm -split "`n"
            $DiffCount = 0
            $MaxLen = [Math]::Max($SrcArr.Count, $WbkArr.Count)
            for ($i = 0; $i -lt $MaxLen; $i++) {
                $s = if ($i -lt $SrcArr.Count) { $SrcArr[$i] } else { "" }
                $w = if ($i -lt $WbkArr.Count) { $WbkArr[$i] } else { "" }
                if ($s -ne $w) { $DiffCount++ }
            }

            $Status = if ($SrcLines -gt $WbkLines) { "SRC_NEWER" } else { "WBK_NEWER" }
            $Results.Add([PSCustomObject]@{
                Name = $Name; Type = $TypeLabel; Status = $Status
                SrcLines = $SrcLines; WbkLines = $WbkLines
                Detail = "$DiffCount differing line(s)"
            })
        }
    }

    # ── Source-only components ────────────────────────────────────────────────
    foreach ($Name in $SourceMap.Keys) {
        if (-not $WorkbookNames.Contains($Name)) {
            $SrcFile = $SourceMap[$Name].SrcPath
            $SrcLines = (Get-Content $SrcFile | Measure-Object -Line).Lines
            $Results.Add([PSCustomObject]@{
                Name = $Name; Type = "?"; Status = "ONLY_IN_SRC"
                SrcLines = $SrcLines; WbkLines = 0
                Detail = "Exists in src\, not in workbook"
            })
        }
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

    if (Test-Path $TempExportDir) {
        Remove-Item $TempExportDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# ─── Build report ─────────────────────────────────────────────────────────────

$Sorted = $Results | Sort-Object Status, Name

$ColWidth = @{ Name = 32; Type = 10; Status = 14; Src = 6; Wbk = 6; Detail = 40 }
$Header   = "{0,-32} {1,-10} {2,-14} {3,6} {4,6}  {5}" -f "Name", "Type", "Status", "Src", "Wbk", "Detail"
$Divider  = "-" * 90

$Lines = @()
$Lines += ""
$Lines += "========================================================================"
$Lines += " XLSmart VBA Comparison Report"
$Lines += " Workbook: $WorkbookPath"
$Lines += " Source:   $SourcePath"
$Lines += " Date:     $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$Lines += "========================================================================"
$Lines += $Header
$Lines += $Divider

$Counts = @{ IDENTICAL = 0; SRC_NEWER = 0; WBK_NEWER = 0; ONLY_IN_SRC = 0; ONLY_IN_WBK = 0; ERROR = 0 }

foreach ($R in $Sorted) {
    $Icon = switch ($R.Status) {
        "IDENTICAL"    { "=" }
        "SRC_NEWER"    { "^" }
        "WBK_NEWER"    { "v" }
        "ONLY_IN_SRC"  { "+" }
        "ONLY_IN_WBK"  { "!" }
        "EXPORT_ERROR" { "X" }
        "UNKNOWN_TYPE" { "?" }
        default        { " " }
    }
    $Row = "{0,-32} {1,-10} {2,-14} {3,6} {4,6}  {5}" -f `
        "$Icon $($R.Name)", $R.Type, $R.Status, $R.SrcLines, $R.WbkLines, $R.Detail
    $Lines += $Row

    $key = if ($Counts.ContainsKey($R.Status)) { $R.Status }
           elseif ($R.Status -match "ERROR") { "ERROR" }
           else { "IDENTICAL" }
    $Counts[$key]++
}

$Lines += $Divider
$Lines += ""
$Lines += "Summary:"
$Lines += "  Identical      : $($Counts.IDENTICAL)"
$Lines += "  Source newer   : $($Counts.SRC_NEWER)   (src has more lines / changed)"
$Lines += "  Workbook newer : $($Counts.WBK_NEWER)   (workbook has more lines / changed)"
$Lines += "  Only in src    : $($Counts.ONLY_IN_SRC)   (not yet imported to workbook)"
$Lines += "  Only in wbk    : $($Counts.ONLY_IN_WBK)  (not exported to src)"
$Lines += "  Errors         : $($Counts.ERROR)"
$Lines += ""
$Lines += "Status legend:"
$Lines += "  = IDENTICAL    No differences detected."
$Lines += "  ^ SRC_NEWER    Local source has more lines or changed content."
$Lines += "  v WBK_NEWER    Workbook copy has more lines or changed content."
$Lines += "  + ONLY_IN_SRC  Component exists in src\ but not in workbook."
$Lines += "  ! ONLY_IN_WBK  Component exists in workbook but not in src\."
$Lines += "  X EXPORT_ERROR Could not export component for comparison."
$Lines += ""
$Lines += "========================================================================"

# Print report
$Lines | ForEach-Object { Write-Host $_ -ForegroundColor White }

# Colour the status column for readability
Write-Host "`n(^ = SRC_NEWER → run Import-VBA.ps1 to sync into workbook)" -ForegroundColor Yellow
Write-Host "(! = ONLY_IN_WBK → run Export-VBA.ps1 to capture into src\)`n" -ForegroundColor Yellow

# Save report
if (-not [string]::IsNullOrEmpty($OutputReport)) {
    $Lines | Set-Content $OutputReport -Encoding UTF8
    Write-Host "Report saved: $OutputReport" -ForegroundColor Green
}

return $Results

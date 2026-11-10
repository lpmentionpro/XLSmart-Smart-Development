$ErrorActionPreference = 'Stop'

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open((Resolve-Path 'workbook/XLSmart_User.xlam').Path, $false, $true)
    $project = $workbook.VBProject

    $engine = $project.VBComponents.Item('modPowerDashboardEngine').CodeModule
    $ribbon = $project.VBComponents.Item('modRibbon').CodeModule
    $form = $project.VBComponents.Item('frmDashboard').CodeModule

    $engineText = $engine.Lines(1, $engine.CountOfLines)
    $ribbonText = $ribbon.Lines(1, $ribbon.CountOfLines)
    $formText = $form.Lines(1, $form.CountOfLines)

    $templatePath = Join-Path (Resolve-Path '.').Path 'assets\PowerDashboardTemplate.html'
    $templateText = if (Test-Path $templatePath) { [System.IO.File]::ReadAllText($templatePath) } else { '' }

    $checks = [ordered]@{
        'Professional dashboard template is bundled' = $templateText.Length -gt 800000
        'Template has data injection hooks' = $templateText.Contains('/*--RAWDATA_START--*/') -and $templateText.Contains('/*--AICONFIG_START--*/')
        'VBA injects Excel data into template' = $engineText.Contains('ReplaceTemplateSection') -and $engineText.Contains('let rawData = ')
        'VBA can cache the template after update' = $engineText.Contains('DownloadDashboardTemplate') -and $engineText.Contains('assets/PowerDashboardTemplate.html')
        'Modeless dashboard form' = $ribbonText.Contains('frmDashboard.Show vbModeless')
        'Native range picker supports keyboard selection' = $formText.Contains('Application.InputBox') -and $formText.Contains('Type:=8') -and $formText.Contains('Ctrl+Shift+Arrow') -and -not $formText.Contains('mSelectingRange')
    }

    $checks.GetEnumerator() | ForEach-Object {
        Write-Host ("{0}: {1}" -f $_.Key, $(if ($_.Value) { 'PASS' } else { 'FAIL' }))
    }

    if ($checks.Values -contains $false) { exit 1 }
}
finally {
    if ($workbook) { $workbook.Close($false); [Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null }
    if ($excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

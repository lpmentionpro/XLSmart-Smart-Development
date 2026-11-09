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

    $checks = [ordered]@{
        'Corrected chart label generator' = $engineText.Contains('let labs=a.map')
        'Corrected CSS color value' = $engineText.Contains('rgba(255,255,255,.68)') -and -not $engineText.Contains('rgba(255,255,268,.68)')
        'Explicit dashboard DOM lookups' = $engineText.Contains("const currencyCodeSel=`$('currencyCodeSel')")
        'Modeless dashboard form' = $ribbonText.Contains('frmDashboard.Show vbModeless')
        'Manual worksheet picker mode' = $formText.Contains('mSelectingRange') -and $formText.Contains('Use Selected Range') -and -not $formText.Contains('Application.InputBox')
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

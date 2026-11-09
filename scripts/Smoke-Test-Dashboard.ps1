<#!
.SYNOPSIS
    Generates a temporary browser dashboard from a minimal Excel data range.

.DESCRIPTION
    Intended as a manual smoke-test helper. The dashboard macro shows a success
    dialog and opens the browser, so the test must be finished interactively.
#>

$ErrorActionPreference = 'Stop'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false

$wb = $xl.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Name = 'DashTest'
$ws.Range('A1:B3').Value2 = @(
    @('Month', 'Sales'),
    @('Jan', 120),
    @('Feb', 240)
)

$addin = $xl.Workbooks.Open((Resolve-Path 'workbook/XLSmart_User.xlam').Path)

Write-Host 'Excel is open with DashTest data. Select A1:B3 and run Power Dashboard manually.'
Write-Host 'Close Excel when the dashboard verification is complete.'
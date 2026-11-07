Attribute VB_Name = "modRibbon"
Option Explicit

Public myRibbon As IRibbonUI

Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set myRibbon = ribbon
End Sub


' Power Dashboard button
Public Sub Ribbon_OpenPowerDashboard(control As IRibbonControl)

    frmDashboard.Show

End Sub


' Smart Chart button
Public Sub Ribbon_OpenSmartChart(control As IRibbonControl)

    frmSmartChartBuilder.Show

End Sub


Public Sub ShowAbout(control As IRibbonControl)

    MsgBox "XLSmart Excel Add-in" & vbNewLine & vbNewLine & _
           "Features:" & vbNewLine & _
           "  - Workbook Index (Sheet Index, Gallery, Back Buttons)" & vbNewLine & _
           "  - AI Tools (Formula Generator, Explainer, Error Detector, Comment Generator)" & vbNewLine & vbNewLine & _
           "Version: 1.0", _
           vbInformation, "About XLSmart"

End Sub


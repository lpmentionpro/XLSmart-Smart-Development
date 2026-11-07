Attribute VB_Name = "modRibbonCallbacks"
Option Explicit

Public Sub ShowAutoFileOrganizer(control As IRibbonControl)

    RunAutoFileOrganizer

End Sub

Public Sub ShowSmartDataCleaner(control As IRibbonControl)

    frmSmartDataCleaner.Show vbModeless

End Sub

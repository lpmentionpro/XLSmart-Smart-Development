VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =========================================================
' THISWORKBOOK - XLSmart
' =========================================================
Dim appEvent As New clsFocuscellWok

Private Sub Workbook_Open()

    On Error Resume Next
    Set appEvent.App = Application
    On Error GoTo 0

End Sub

Public Sub ActivateAppEvent()

    Set appEvent.App = Application

End Sub

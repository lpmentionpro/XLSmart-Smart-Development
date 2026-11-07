VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFocuscellWok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =========================================================
' CLASS MODULE: clsFocusCellWok
' =========================================================
' Ei class ta Application level event handle kore.
' Mane je kono workbook e cell change hole
' automatically PaintFocusCell call hobe.
' Shudhu xlam er ThisWorkbook theke activate kora hoy.
' =========================================================

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_SheetSelectionChange(ByVal sh As Object, ByVal Target As Range)
    ' Jodi Focus Cell ON thake, tahole
    ' cell change hole highlight move korbe
    If focusIsOn Then
        PaintFocusCell
    End If
End Sub

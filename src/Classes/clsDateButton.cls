VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDateButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Btn As _
    MSForms.CommandButton
Attribute Btn.VB_VarHelpID = -1

Public ParentForm As Object


Private Sub Btn_Click()

    If Btn.Caption = "" _
    Then Exit Sub

    ParentForm.SelectDate _
        CLng(Btn.Caption)

End Sub


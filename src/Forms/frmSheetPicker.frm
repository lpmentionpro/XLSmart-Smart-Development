VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSheetPicker 
   Caption         =   "Pro Data Modeler"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6360
   OleObjectBlob   =   "frmSheetPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSheetPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SelectedSheetIndexes As String
Public headerRow As Long
Public WasCancelled As Boolean

Private Sub UserForm_Initialize()

    Dim i As Long

    WasCancelled = True
    SelectedSheetIndexes = ""
    headerRow = 1

    lstSheets.Clear

    For i = 1 To ActiveWorkbook.Worksheets.count
        lstSheets.AddItem ActiveWorkbook.Worksheets(i).Name
    Next i

    txtHeaderRow.text = "1"

End Sub

Private Sub btnOK_Click()

    Dim i As Long
    Dim result As String

    For i = 0 To lstSheets.ListCount - 1
        If lstSheets.Selected(i) Then
            If result <> "" Then result = result & ","
            result = result & CStr(i + 1)
        End If
    Next i

    If result = "" Then
        MsgBox "Please tick at least one sheet.", vbExclamation
        Exit Sub
    End If

    If val(txtHeaderRow.text) < 1 Then
        MsgBox "Header row must be 1 or more.", vbExclamation
        Exit Sub
    End If

    SelectedSheetIndexes = result
    headerRow = CLng(val(txtHeaderRow.text))
    WasCancelled = False

    Me.Hide

End Sub

Private Sub btnCancel_Click()
    WasCancelled = True
    Me.Hide
End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDateFormatStudio 
   Caption         =   "Date Format Studio"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6360
   OleObjectBlob   =   "frmDateFormatStudio.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDateFormatStudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Initialize()

    Me.Caption = "Date Format Studio"

    cmbDateFormat.Clear

    cmbDateFormat.AddItem "dd-mm-yyyy"
    cmbDateFormat.AddItem "mm-dd-yyyy"
    cmbDateFormat.AddItem "yyyy-mm-dd"
    cmbDateFormat.AddItem "dd/mm/yyyy"
    cmbDateFormat.AddItem "mm/dd/yyyy"
    cmbDateFormat.AddItem "dd mmm yyyy"
    cmbDateFormat.AddItem "dd mmmm yyyy"
    cmbDateFormat.AddItem "dd-mmm-yyyy"
    cmbDateFormat.AddItem "mmm dd, yyyy"
    cmbDateFormat.AddItem "dddd, dd mmmm yyyy"

    cmbDateFormat.ListIndex = 0

    chkAutoApply.Value = True

    UpdatePreview

End Sub


Private Sub cmbDateFormat_Change()

    UpdatePreview

    If chkAutoApply.Value Then
        ApplyDateFormat
    End If

End Sub


Private Sub btnApply_Click()

    ApplyDateFormat

    Unload Me

End Sub


Private Sub btnCancel_Click()

    Unload Me

End Sub


Private Sub UpdatePreview()

    lblPreview.Caption = _
        Format(Date, cmbDateFormat.Value)

End Sub


Private Sub ApplyDateFormat()

    Dim workRng As Range
    Dim c As Range
    Dim lastRow As Long
    Dim dt As Date

    If TypeName(Selection) <> "Range" Then Exit Sub


    If Selection.Columns.count = 1 And _
       Selection.Rows.count = Rows.count Then

        lastRow = Cells(Rows.count, _
                    Selection.Column).End(xlUp).Row

        Set workRng = Range( _
            Cells(1, Selection.Column), _
            Cells(lastRow, Selection.Column))
    Else

        Set workRng = Selection

    End If


    On Error Resume Next

    For Each c In workRng.Cells

        If Trim(c.Value) <> "" Then

            dt = CDate(c.Value)

            If Err.Number = 0 Then

                c.Value = dt

                c.NumberFormat = _
                    cmbDateFormat.Value

                c.HorizontalAlignment = _
                    xlLeft

            End If

            Err.Clear

        End If

    Next c
    
    workRng.EntireColumn.AutoFit

    On Error GoTo 0

End Sub


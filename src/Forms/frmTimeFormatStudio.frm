VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTimeFormatStudio 
   Caption         =   "Date Format Studio"
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6255
   OleObjectBlob   =   "frmTimeFormatStudio.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTimeFormatStudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private targetRange As Range


Private Sub UserForm_Initialize()

    Me.Caption = _
        "Time Format Studio"

    If TypeName(Selection) _
        = "Range" Then

        Set targetRange = _
            Selection

    End If

    chkAutoApply.Value = True
    opt12Hour.Value = True

    LoadTimeFormats
    UpdatePreview

End Sub


'==============================
' LOAD FORMATS
'==============================

Private Sub LoadTimeFormats()

    cmbTimeFormat.Clear

    If opt12Hour.Value Then

        cmbTimeFormat.AddItem _
            "hh:mm AM/PM"

        cmbTimeFormat.AddItem _
            "hh:mm:ss AM/PM"

        cmbTimeFormat.ListIndex = 0

    Else

        cmbTimeFormat.AddItem _
            "HH:mm"

        cmbTimeFormat.AddItem _
            "HH:mm:ss"

        cmbTimeFormat.ListIndex = 0

    End If

End Sub


'==============================
' FORMAT OPTION
'==============================

Private Sub opt12Hour_Click()

    LoadTimeFormats
    UpdatePreview

    If chkAutoApply.Value Then
        ApplyTimeFormat
    End If

End Sub


Private Sub opt24Hour_Click()

    LoadTimeFormats
    UpdatePreview

    If chkAutoApply.Value Then
        ApplyTimeFormat
    End If

End Sub


Private Sub cmbTimeFormat_Change()

    UpdatePreview

    If chkAutoApply.Value Then
        ApplyTimeFormat
    End If

End Sub


Private Sub chkAutoApply_Click()

    If chkAutoApply.Value Then
        ApplyTimeFormat
    End If

End Sub


'==============================
' PREVIEW
'==============================

Private Sub UpdatePreview()

    Dim v As Variant
    Dim tm As Date

    If targetRange _
        Is Nothing Then Exit Sub

    v = _
        targetRange.Cells(1, 1).Value

    If IsEmpty(v) Then

        lblPreview.Caption = _
            "No Time"

        Exit Sub

    End If


    On Error Resume Next

    ' Excel serial time support
    tm = TimeValue(CDate(v))

    If Err.Number <> 0 Then

        lblPreview.Caption = _
            "Invalid Time"

        Err.Clear

    Else

        lblPreview.Caption = _
            Format( _
            tm, _
            cmbTimeFormat.Value)

    End If

    On Error GoTo 0

End Sub


'==============================
' APPLY FORMAT
'==============================

Private Sub ApplyTimeFormat()

    Dim c As Range
    Dim tm As Date
    Dim v As Variant

    If targetRange _
        Is Nothing Then Exit Sub

    Application.ScreenUpdating = _
        False

    For Each c In _
        targetRange.Cells

        If Not IsEmpty(c.Value) Then

            v = c.Value

            On Error Resume Next

            tm = _
                TimeValue(CDate(v))

            If Err.Number = 0 Then

                c.Value = tm

                c.NumberFormat = _
                    cmbTimeFormat.Value

            End If

            Err.Clear
            On Error GoTo 0

        End If

    Next c

    Application.ScreenUpdating = _
        True

    UpdatePreview

End Sub


'==============================
' BUTTONS
'==============================
Private Sub btnApply_Click()

    ApplyTimeFormat

    Unload Me

End Sub


Private Sub btnCancel_Click()

    Unload Me

End Sub


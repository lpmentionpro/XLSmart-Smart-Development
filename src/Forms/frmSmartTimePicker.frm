VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSmartTimePicker 
   Caption         =   "Smart Time Picker"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6165
   OleObjectBlob   =   "frmSmartTimePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSmartTimePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private selectedTime As Date


Private Sub UserForm_Initialize()

    Dim i As Long

    ' Hour
    cmbHour.Clear

    For i = 0 To 23
        cmbHour.AddItem _
            Format(i, "00")
    Next i


    ' Minute
    cmbMinute.Clear

    For i = 0 To 59
        cmbMinute.AddItem _
            Format(i, "00")
    Next i


    ' Meridian
    cmbMeridian.Clear
    cmbMeridian.AddItem "AM"
    cmbMeridian.AddItem "PM"


    ' Default format
    opt12Hour.Value = True
    opt24Hour.Value = False


    ' Current time
    LoadCurrentTime

    UpdatePreview

End Sub


'==================================
' LOAD CURRENT TIME
'==================================

Private Sub LoadCurrentTime()

    Dim h As Integer

    h = Hour(Time)

    If opt12Hour.Value Then

        If h = 0 Then
            h = 12
        ElseIf h > 12 Then
            h = h - 12
        End If

        cmbHour.Value = _
            Format(h, "00")

        cmbMeridian.Visible = True

        If Hour(Time) >= 12 Then
            cmbMeridian.Value = "PM"
        Else
            cmbMeridian.Value = "AM"
        End If

    Else

        cmbHour.Value = _
            Format(Hour(Time), "00")

        cmbMeridian.Visible = False

    End If


    cmbMinute.Value = _
        Format(Minute(Time), "00")

End Sub


'==================================
' FORMAT SWITCH
'==================================

Private Sub opt12Hour_Click()

    cmbHour.Clear

    Dim i As Long

    For i = 1 To 12
        cmbHour.AddItem _
            Format(i, "00")
    Next i

    cmbMeridian.Visible = True

    LoadCurrentTime
    UpdatePreview

End Sub


Private Sub opt24Hour_Click()

    cmbHour.Clear

    Dim i As Long

    For i = 0 To 23
        cmbHour.AddItem _
            Format(i, "00")
    Next i

    cmbMeridian.Visible = False

    LoadCurrentTime
    UpdatePreview

End Sub


'==================================
' LIVE PREVIEW
'==================================

Private Sub cmbHour_Change()

    UpdatePreview

End Sub


Private Sub cmbMinute_Change()

    UpdatePreview

End Sub


Private Sub cmbMeridian_Change()

    UpdatePreview

End Sub


Private Sub UpdatePreview()

    Dim h As Integer
    Dim m As Integer
    Dim mer As String

    If cmbHour.Value = "" Then Exit Sub
    If cmbMinute.Value = "" Then Exit Sub


    h = val(cmbHour.Value)
    m = val(cmbMinute.Value)

    If opt12Hour.Value Then

        mer = cmbMeridian.Value

        If mer = "PM" _
        And h <> 12 Then

            h = h + 12

        End If

        If mer = "AM" _
        And h = 12 Then

            h = 0

        End If

    End If


    selectedTime = _
        TimeSerial(h, m, 0)


    If opt12Hour.Value Then

        lblPreview.Caption = _
            Format(selectedTime, _
            "hh:mm AM/PM")

    Else

        lblPreview.Caption = _
            Format(selectedTime, _
            "HH:mm")

    End If

End Sub


'==================================
' NOW BUTTON
'==================================

Private Sub btnNow_Click()

    LoadCurrentTime
    UpdatePreview

End Sub


'==================================
' APPLY
'==================================

Private Sub btnApply_Click()

    Dim workRng As Range
    Dim c As Range
    Dim lastRow As Long

    If TypeName(Selection) <> _
        "Range" Then Exit Sub


    If Selection.Columns.count = 1 _
    And Selection.Rows.count = _
    Rows.count Then

        lastRow = Cells( _
            Rows.count, _
            Selection.Column). _
            End(xlUp).Row

        Set workRng = _
            Range( _
            Cells(1, _
            Selection.Column), _
            Cells(lastRow, _
            Selection.Column))

    Else

        Set workRng = Selection

    End If


    For Each c In _
        workRng.Cells

        c.Value = selectedTime

        If opt12Hour.Value Then

            c.NumberFormat = _
                "hh:mm AM/PM"

        Else

            c.NumberFormat = _
                "hh:mm"

        End If

        c.HorizontalAlignment = _
            xlLeft

    Next c


    workRng.EntireColumn.AutoFit

    Unload Me

End Sub


Private Sub btnCancel_Click()

    Unload Me

End Sub


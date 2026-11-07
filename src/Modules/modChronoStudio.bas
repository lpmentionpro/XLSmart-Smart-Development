Attribute VB_Name = "modChronoStudio"
Option Explicit



'==================================
' PICKER ACTIONS
'==================================

Public Sub Chrono_PickDate( _
    control As IRibbonControl)

    frmSmartDatePicker.Show

End Sub


Public Sub Chrono_PickTime( _
    control As IRibbonControl)

    frmSmartTimePicker.Show

End Sub


Public Sub Chrono_PickDateTime( _
    control As IRibbonControl)

    frmSmartDateTimePicker.Show

End Sub



'==================================
' QUICK INSERT
'==================================

Public Sub Chrono_InsertTodayDate( _
    control As IRibbonControl)

    Dim workRng As Range

    If TypeName(Selection) <> _
        "Range" Then Exit Sub

    Set workRng = GetSmartRange

    workRng.Value = Date
    workRng.NumberFormat = _
        "dd-mm-yyyy"

    workRng.HorizontalAlignment = _
        xlLeft

    workRng.EntireColumn.AutoFit

End Sub


Public Sub Chrono_InsertCurrentDateTime( _
    control As IRibbonControl)

    Dim workRng As Range

    If TypeName(Selection) <> _
        "Range" Then Exit Sub

    Set workRng = GetSmartRange

    workRng.Value = Now

    workRng.NumberFormat = _
        "dd-mm-yyyy hh:mm:ss"

    workRng.HorizontalAlignment = _
        xlLeft

    workRng.EntireColumn.AutoFit

End Sub


Public Sub Chrono_InsertCurrentTime( _
    control As IRibbonControl)

    Dim workRng As Range

    If TypeName(Selection) <> _
        "Range" Then Exit Sub

    Set workRng = GetSmartRange

    workRng.Value = Time

    workRng.NumberFormat = _
        "hh:mm:ss"

    workRng.HorizontalAlignment = _
        xlLeft

    workRng.EntireColumn.AutoFit

End Sub


'==================================
' SETTINGS
'==================================

Public Sub Chrono_OpenDateSettings( _
    control As IRibbonControl)

    frmDateFormatStudio.Show

End Sub


Public Sub Chrono_OpenTimeSettings( _
    control As IRibbonControl)

    frmTimeFormatStudio.Show

End Sub


'==================================
' SMART RANGE
'==================================

Private Function GetSmartRange() _
    As Range

    Dim lastRow As Long

    If Selection.Columns.count = 1 _
    And Selection.Rows.count = _
    Rows.count Then

        lastRow = Cells( _
            Rows.count, _
            Selection.Column). _
            End(xlUp).Row

        Set GetSmartRange = _
            Range( _
            Cells(1, _
            Selection.Column), _
            Cells(lastRow, _
            Selection.Column))

    Else

        Set GetSmartRange = _
            Selection

    End If

End Function


Public Sub Chrono_LoadSettings()

End Sub


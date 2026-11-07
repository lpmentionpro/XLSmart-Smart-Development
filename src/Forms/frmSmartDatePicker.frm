VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSmartDatePicker 
   Caption         =   "Smart Date Picker"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6960
   OleObjectBlob   =   "frmSmartDatePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSmartDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private currentMonth As Integer
Private currentYear As Integer
Private selectedDate As Date
Private btnEvents As Collection


Private Sub UserForm_Initialize()

    Dim i As Integer

    cmbMonth.Clear

    cmbMonth.AddItem "January"
    cmbMonth.AddItem "February"
    cmbMonth.AddItem "March"
    cmbMonth.AddItem "April"
    cmbMonth.AddItem "May"
    cmbMonth.AddItem "June"
    cmbMonth.AddItem "July"
    cmbMonth.AddItem "August"
    cmbMonth.AddItem "September"
    cmbMonth.AddItem "October"
    cmbMonth.AddItem "November"
    cmbMonth.AddItem "December"


    cmbYear.Clear

    For i = Year(Date) - 20 To Year(Date) + 20
        cmbYear.AddItem i
    Next i


    currentMonth = Month(Date)
    currentYear = Year(Date)

    cmbMonth.ListIndex = currentMonth - 1
    cmbYear.Value = currentYear

    chkSequenceFill.Value = True
    chkSameDate.Value = False

    selectedDate = Date
    InitButtonEvents
    RenderCalendar
    UpdatePreview

End Sub


'========================================
' CALENDAR RENDER
'========================================

Private Sub RenderCalendar()

    Dim firstDay As Date
    Dim startPos As Integer
    Dim daysInMonth As Integer
    Dim i As Integer
    Dim Btn As MSForms.CommandButton
    Dim d As Integer

    firstDay = DateSerial(currentYear, currentMonth, 1)

    startPos = Weekday(firstDay, vbSunday)

    daysInMonth = Day( _
        DateSerial(currentYear, _
        currentMonth + 1, 0))


    ' clear all buttons
    For i = 1 To 42

        Set Btn = Me.Controls( _
            "btnDay" & i)

        Btn.Caption = ""
        Btn.Enabled = False
        Btn.BackColor = RGB(240, 240, 240)

    Next i


    d = 1

    For i = startPos To _
        startPos + daysInMonth - 1

        Set Btn = Me.Controls( _
            "btnDay" & i)

        Btn.Caption = d
        Btn.Enabled = True

        If DateSerial( _
            currentYear, _
            currentMonth, d) = selectedDate Then

            Btn.BackColor = _
                RGB(80, 170, 80)

        Else

            Btn.BackColor = _
                RGB(255, 255, 255)

        End If

        d = d + 1

    Next i

End Sub


'========================================
' PREVIEW
'========================================

Private Sub UpdatePreview()

    lblPreview.Caption = _
        Format(selectedDate, _
        "dddd, dd mmmm yyyy")

End Sub


'========================================
' MONTH / YEAR
'========================================

Private Sub cmbMonth_Change()

    currentMonth = _
        cmbMonth.ListIndex + 1

    RenderCalendar

End Sub


Private Sub cmbYear_Change()

    If cmbYear.Value = "" _
    Then Exit Sub

    currentYear = _
        CLng(cmbYear.Value)

    RenderCalendar

End Sub


Private Sub btnPrevMonth_Click()

    currentMonth = _
        currentMonth - 1

    If currentMonth < 1 Then

        currentMonth = 12
        currentYear = _
            currentYear - 1

    End If

    cmbMonth.ListIndex = _
        currentMonth - 1

    cmbYear.Value = _
        currentYear

    RenderCalendar

End Sub


Private Sub btnNextMonth_Click()

    currentMonth = _
        currentMonth + 1

    If currentMonth > 12 Then

        currentMonth = 1
        currentYear = _
            currentYear + 1

    End If

    cmbMonth.ListIndex = _
        currentMonth - 1

    cmbYear.Value = _
        currentYear

    RenderCalendar

End Sub


'========================================
' TODAY
'========================================

Private Sub btnToday_Click()

    selectedDate = Date

    currentMonth = _
        Month(Date)

    currentYear = _
        Year(Date)

    cmbMonth.ListIndex = _
        currentMonth - 1

    cmbYear.Value = _
        currentYear

    RenderCalendar
    UpdatePreview

End Sub


'========================================
' MODE CHECKBOX
'========================================

Private Sub chkSequenceFill_Click()

    If chkSequenceFill.Value _
    Then

        chkSameDate.Value = _
            False

    End If

End Sub


Private Sub chkSameDate_Click()

    If chkSameDate.Value _
    Then

        chkSequenceFill.Value = _
            False

    End If

End Sub


'========================================
' APPLY
'========================================

Private Sub btnApply_Click()

    Dim workRng As Range
    Dim c As Range
    Dim i As Long
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

        Set workRng = _
            Selection

    End If


    i = 0

    For Each c In _
        workRng.Cells

        If chkSequenceFill.Value _
        Then

            c.Value = _
                selectedDate + i

        Else

            c.Value = _
                selectedDate

        End If

        c.NumberFormat = _
            "dd-mmm-yyyy"

        c.HorizontalAlignment = _
            xlLeft

        i = i + 1

    Next c


    workRng.EntireColumn.AutoFit

    Unload Me

End Sub


Private Sub btnCancel_Click()

    Unload Me

End Sub


'========================================
' BUTTON EVENTS
'========================================

Private Sub InitButtonEvents()

    Dim i As Long
    Dim btnObj As clsDateButton

    Set btnEvents = _
        New Collection

    For i = 1 To 42

        Set btnObj = _
            New clsDateButton

        Set btnObj.Btn = _
            Me.Controls( _
            "btnDay" & i)

        Set btnObj.ParentForm = _
            Me

        btnEvents.Add btnObj

    Next i

End Sub


Public Sub SelectDate( _
    dayNumber As Long)

    selectedDate = _
        DateSerial( _
        currentYear, _
        currentMonth, _
        dayNumber)

    RenderCalendar
    UpdatePreview

End Sub


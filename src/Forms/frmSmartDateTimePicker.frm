VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSmartDateTimePicker 
   Caption         =   "Smart Date & Time Picker "
   ClientHeight    =   9810.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7560
   OleObjectBlob   =   "frmSmartDateTimePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSmartDateTimePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private currentMonth As Integer
Private currentYear As Integer

Private selectedDate As Date
Private selectedTime As Date

Private btnEvents As Collection


'====================================
' FORM LOAD
'====================================

Private Sub UserForm_Initialize()

    Dim i As Long

    '-------------------------
    ' MONTH
    '-------------------------
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


    '-------------------------
    ' YEAR
    '-------------------------
    cmbYear.Clear

    For i = Year(Date) - 20 _
        To Year(Date) + 20

        cmbYear.AddItem i

    Next i


    currentMonth = Month(Date)
    currentYear = Year(Date)

    cmbMonth.ListIndex = _
        currentMonth - 1

    cmbYear.Value = _
        currentYear


    '-------------------------
    ' MINUTES
    '-------------------------
    Me.fraTime.Controls( _
        "cmbMinute").Clear

    For i = 0 To 59

        Me.fraTime.Controls( _
            "cmbMinute").AddItem _
            Format(i, "00")

    Next i


    '-------------------------
    ' MERIDIAN
    '-------------------------
    Me.fraTime.Controls( _
        "cmbMeridian").Clear

    Me.fraTime.Controls( _
        "cmbMeridian").AddItem "AM"

    Me.fraTime.Controls( _
        "cmbMeridian").AddItem "PM"


    opt12Hour.Value = True

    Load12HourList
    LoadCurrentTime


    selectedDate = Date

    InitButtonEvents
    RenderCalendar
    UpdatePreview

End Sub


'====================================
' DATE BUTTON EVENTS
'====================================

Private Sub InitButtonEvents()

    Dim i As Long
    Dim btnObj As clsDateButton

    Set btnEvents = _
        New Collection

    For i = 1 To 42

        Set btnObj = _
            New clsDateButton

        Set btnObj.Btn = _
            Me.Frame1.Controls( _
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


'====================================
' CALENDAR
'====================================

Private Sub RenderCalendar()

    Dim firstDay As Date
    Dim startPos As Integer
    Dim daysInMonth As Integer

    Dim i As Long
    Dim d As Long

    Dim Btn As MSForms.CommandButton

    firstDay = _
        DateSerial( _
        currentYear, _
        currentMonth, 1)

    startPos = _
        Weekday(firstDay, _
        vbSunday)

    daysInMonth = _
        Day(DateSerial( _
        currentYear, _
        currentMonth + 1, 0))


    For i = 1 To 42

        Set Btn = _
            Me.Frame1.Controls( _
            "btnDay" & i)

        Btn.Caption = ""
        Btn.Enabled = False

        Btn.BackColor = _
            RGB(255, 255, 255)

    Next i


    d = 1

    For i = startPos _
        To startPos + _
        daysInMonth - 1

        Set Btn = _
            Me.Frame1.Controls( _
            "btnDay" & i)

        Btn.Caption = d
        Btn.Enabled = True

        If DateSerial( _
            currentYear, _
            currentMonth, d) _
            = selectedDate Then

            Btn.BackColor = _
                RGB(70, 180, 90)

        End If

        d = d + 1

    Next i

End Sub


'====================================
' MONTH / YEAR
'====================================

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


'====================================
' LOAD TIME LIST
'====================================

Private Sub Load12HourList()

    Dim i As Long

    Me.fraTime.Controls( _
        "cmbHour").Clear

    For i = 1 To 12

        Me.fraTime.Controls( _
            "cmbHour").AddItem _
            Format(i, "00")

    Next i

End Sub


Private Sub Load24HourList()

    Dim i As Long

    Me.fraTime.Controls( _
        "cmbHour").Clear

    For i = 0 To 23

        Me.fraTime.Controls( _
            "cmbHour").AddItem _
            Format(i, "00")

    Next i

End Sub


Private Sub LoadCurrentTime()

    Dim h As Integer

    h = Hour(Time)

    If opt12Hour.Value Then

        If h = 0 Then
            h = 12
        ElseIf h > 12 Then
            h = h - 12
        End If

        Me.fraTime.Controls( _
            "cmbHour").Value = _
            Format(h, "00")

        Me.fraTime.Controls( _
            "cmbMeridian").Visible = True

        If Hour(Time) >= 12 Then

            Me.fraTime.Controls( _
                "cmbMeridian").Value = _
                "PM"

        Else

            Me.fraTime.Controls( _
                "cmbMeridian").Value = _
                "AM"

        End If

    Else

        Me.fraTime.Controls( _
            "cmbHour").Value = _
            Format(Hour(Time), "00")

        Me.fraTime.Controls( _
            "cmbMeridian").Visible = False

    End If


    Me.fraTime.Controls( _
        "cmbMinute").Value = _
        Format(Minute(Time), "00")

End Sub


'====================================
' TIME EVENTS
'====================================

Private Sub opt12Hour_Click()

    Load12HourList
    LoadCurrentTime
    UpdatePreview

End Sub


Private Sub opt24Hour_Click()

    Load24HourList
    LoadCurrentTime
    UpdatePreview

End Sub


Private Sub cmbHour_Change()

    UpdatePreview

End Sub


Private Sub cmbMinute_Change()

    UpdatePreview

End Sub


Private Sub cmbMeridian_Change()

    UpdatePreview

End Sub


'====================================
' PREVIEW
'====================================

Private Sub UpdatePreview()

    Dim h As Integer
    Dim m As Integer

    If Me.fraTime.Controls( _
        "cmbHour").Value = "" _
    Then Exit Sub

    If Me.fraTime.Controls( _
        "cmbMinute").Value = "" _
    Then Exit Sub


    h = val(Me.fraTime.Controls( _
        "cmbHour").Value)

    m = val(Me.fraTime.Controls( _
        "cmbMinute").Value)


    If opt12Hour.Value Then

        If Me.fraTime.Controls( _
            "cmbMeridian").Value _
            = "PM" _
        And h <> 12 Then

            h = h + 12

        End If


        If Me.fraTime.Controls( _
            "cmbMeridian").Value _
            = "AM" _
        And h = 12 Then

            h = 0

        End If

    End If


    selectedTime = _
        TimeSerial(h, m, 0)


    If opt12Hour.Value Then

        lblPreview.Caption = _
            Format(selectedDate, _
            "dd-mmm-yyyy ") _
            & Format( _
            selectedTime, _
            "hh:mm AM/PM")

    Else

        lblPreview.Caption = _
            Format(selectedDate, _
            "dd-mmm-yyyy ") _
            & Format( _
            selectedTime, _
            "HH:mm")

    End If

End Sub


'====================================
' NOW BUTTON
'====================================

Private Sub btnNow_Click()

    selectedDate = Date

    currentMonth = _
        Month(Date)

    currentYear = _
        Year(Date)

    cmbMonth.ListIndex = _
        currentMonth - 1

    cmbYear.Value = _
        currentYear

    LoadCurrentTime

    RenderCalendar
    UpdatePreview

End Sub


'====================================
' APPLY
'====================================

Private Sub btnApply_Click()

    Dim c As Range
    Dim workRng As Range
    Dim finalDateTime As Date

    If TypeName(Selection) _
        <> "Range" _
    Then Exit Sub


    finalDateTime = _
        selectedDate + _
        selectedTime


    Set workRng = Selection


    For Each c In _
        workRng.Cells

        c.Value = _
            finalDateTime

        If opt12Hour.Value Then

            c.NumberFormat = _
                "dd-mmm-yyyy hh:mm AM/PM"

        Else

            c.NumberFormat = _
                "dd-mmm-yyyy HH:mm"

        End If

    Next c


    workRng.EntireColumn.AutoFit

    Unload Me

End Sub


Private Sub btnCancel_Click()

    Unload Me

End Sub



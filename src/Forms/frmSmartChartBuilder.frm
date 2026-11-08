VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSmartChartBuilder 
   Caption         =   "Smart Chart Builder"
   ClientHeight    =   9810.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14160
   OleObjectBlob   =   "frmSmartChartBuilder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSmartChartBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pickedRange As Range

Private Sub UserForm_Initialize()

    ' chart types
    With cmbChartType
        .AddItem "Column Chart"
        .AddItem "Bar Chart"
        .AddItem "Line Chart"
        .AddItem "Pie Chart"
        .AddItem "Doughnut Chart"
        .AddItem "Area Chart"
        .AddItem "Scatter Chart"
    End With

    ' themes
    With cmbTheme
        .AddItem "Light"
        .AddItem "Dark"
        .AddItem "Ocean"
        .AddItem "Midnight"
        .AddItem "Corporate"
    End With

    ' currency
    With cmbCurrency
        .AddItem "None"
        .AddItem "BDT (Tk)"
        .AddItem "$ USD"
        .AddItem "€ EUR"
        .AddItem "? INR"
    End With

    ' grid line
    With cmbGrid
        .AddItem "Show"
        .AddItem "Hide"
    End With

    ' default values
    cmbChartType.ListIndex = 0
    cmbTheme.ListIndex = 0
    cmbCurrency.ListIndex = 1
    cmbGrid.ListIndex = 0

    txtChartTitle = "Sales Chart"
    txtXAxis = "Category"
    txtYAxis = "Value"

    lblStatus.Caption = "Status: Ready"

End Sub


Private Sub btnPickData_Click()

    On Error Resume Next

    Set pickedRange = Application.InputBox( _
        prompt:="Select chart data range", _
        Title:="Pick Data", _
        Type:=8)

    If pickedRange Is Nothing Then Exit Sub

    txtDataRange.text = _
        pickedRange.Worksheet.Name & "!" & pickedRange.Address

    lblStatus.Caption = _
        "Status: " & pickedRange.Rows.count & _
        " rows selected"

End Sub


Private Sub btnLoadDemo_Click()

    txtChartTitle = "Monthly Sales"
    txtXAxis = "Month"
    txtYAxis = "Sales"

    cmbChartType.Value = "Column Chart"
    cmbTheme.Value = "Ocean"
    cmbCurrency.Value = "? BDT"
    cmbGrid.Value = "Show"

    lblStatus.Caption = "Status: Demo loaded"

End Sub


Private Sub btnCancel_Click()

    Unload Me

End Sub


Private Sub btnCreateChart_Click()

    Dim result As VbMsgBoxResult

    If pickedRange Is Nothing Then

        MsgBox "Please select chart data first.", vbExclamation
        Exit Sub

    End If

    result = MsgBox( _
        "Create Interactive Web Chart?" & vbCrLf & vbCrLf & _
        "Yes = HTML Web Chart" & vbCrLf & _
        "No = Excel Chart", _
        vbYesNoCancel + vbQuestion, _
        "Select Output")

    If result = vbCancel Then Exit Sub

    lblStatus.Caption = "Status: Creating chart..."

    ' YES = web chart
    If result = vbYes Then

    Call CreateWebChart( _
        pickedRange, _
        cmbChartType.Value, _
        txtChartTitle.text, _
        cmbTheme.Value, _
        txtXAxis.text, _
        txtYAxis.text)

    Else

    ' Excel chart
    CreateExcelChart

    End If

    lblStatus.Caption = "Status: Done"

End Sub

Private Sub CreateExcelChart()

    Dim ws As Worksheet
    Dim rng As Range
    Dim chartObj As ChartObject
    Dim chartKind As XlChartType
    Dim chartFontColor As Long
    Dim gridColor As Long

    Dim rangeText As String
    Dim sheetName As String
    Dim rangeAddress As String
    Dim splitPos As Long

    rangeText = txtDataRange.text

    splitPos = InStr(rangeText, "!")

    sheetName = left(rangeText, splitPos - 1)
    rangeAddress = Mid(rangeText, splitPos + 1)

    Set ws = Worksheets(sheetName)
    Set rng = ws.Range(rangeAddress)


    Select Case cmbChartType.Value

        Case "Column Chart"
            chartKind = xlColumnClustered

        Case "Bar Chart"
            chartKind = xlBarClustered

        Case "Line Chart"
            chartKind = xlLine

        Case "Pie Chart"
            chartKind = xlPie

        Case "Doughnut Chart"
            chartKind = xlDoughnut

        Case "Area Chart"
            chartKind = xlArea

        Case Else
            chartKind = xlColumnClustered

    End Select

    Set chartObj = ws.ChartObjects.Add( _
        left:=450, _
        Width:=520, _
        Top:=100, _
        Height:=300)

    With chartObj.Chart

    .chartType = chartKind

    ' clear default series
    Do While .SeriesCollection.count > 0
        .SeriesCollection(1).Delete
    Loop

    ' first column = label
    .SeriesCollection.NewSeries
    .SeriesCollection(1).XValues = rng.Columns(1)

    ' second column = value
    .SeriesCollection(1).values = rng.Columns(2)

    ' header name
    .SeriesCollection(1).Name = rng.Cells(1, 2).Value

    ' title
    .HasTitle = True
    .chartTitle.text = txtChartTitle.text
    .chartTitle.Format.TextFrame2.TextRange.Font.Size = 18
    .chartTitle.Format.TextFrame2.TextRange.Font.Bold = True

    ' axis title
    If chartKind <> xlPie And chartKind <> xlDoughnut Then

        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.text = txtXAxis.text

        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.text = txtYAxis.text

    End If

    ' legend
    .HasLegend = False

    ' smooth chart border
    .ChartArea.Format.Line.Visible = msoFalse
    .PlotArea.Format.Line.Visible = msoFalse

    ' grid line
    If cmbGrid.Value = "Hide" Then
        On Error Resume Next
        .Axes(xlValue).MajorGridlines.Delete
        On Error GoTo 0
    Else
        On Error Resume Next
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(220, 220, 220)
        On Error GoTo 0
    End If

    ' ========= THEME =========

    Select Case cmbTheme.Value

    Case "Light"

        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(248, 249, 250)

        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(59, 130, 246)

        chartFontColor = RGB(30, 30, 30)
        gridColor = RGB(220, 220, 220)

    Case "Dark"

        .ChartArea.Format.Fill.ForeColor.RGB = RGB(24, 24, 24)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(38, 38, 38)

        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(0, 200, 255)

        chartFontColor = RGB(255, 255, 255)
        gridColor = RGB(90, 90, 90)

    Case "Ocean"

        .ChartArea.Format.Fill.ForeColor.RGB = RGB(230, 244, 255)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(215, 235, 255)

        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(0, 119, 182)

        chartFontColor = RGB(20, 50, 80)
        gridColor = RGB(180, 210, 230)

    Case "Midnight"

        .ChartArea.Format.Fill.ForeColor.RGB = RGB(18, 18, 35)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(25, 25, 50)

        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(123, 97, 255)

        chartFontColor = RGB(255, 255, 255)
        gridColor = RGB(80, 80, 120)

    Case "Corporate"

        .ChartArea.Format.Fill.ForeColor.RGB = RGB(248, 249, 250)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(240, 242, 245)

        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(0, 91, 150)

        chartFontColor = RGB(35, 35, 35)
        gridColor = RGB(200, 205, 210)

End Select

    ' auto font color by theme
    .chartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = chartFontColor

    On Error Resume Next

    .Axes(xlCategory).TickLabels.Font.Color = chartFontColor
    .Axes(xlValue).TickLabels.Font.Color = chartFontColor

    .Axes(xlCategory).AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = chartFontColor
    .Axes(xlValue).AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = chartFontColor

    If cmbGrid.Value = "Show" Then
    .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = gridColor
    End If

    On Error GoTo 0



    ' rounded style feel
    .ChartArea.Format.Shadow.Visible = msoTrue

End With

    MsgBox "Excel chart created!", vbInformation

End Sub


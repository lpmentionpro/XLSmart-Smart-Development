Attribute VB_Name = "modSmartDataManager"

'=======================================================
' MODULE: modSmartDataManager
' Excel Booster - Data Manager Tools VBA Code
' Tools: Smart Data Manager, Smart Data Exporter,
'        Smart Filter, Smart Highlighter, Smart Masking
'=======================================================
Option Explicit

'-------------------------------------------------------
' GLOBAL VARIABLES
'-------------------------------------------------------
Public g_DataRange As Range
Public g_FilterSheet As Worksheet
Public g_HighlightSheet As Worksheet

Public g_FilterColumn As String
Public g_FilterValue As String

Public gSearchBox As OLEObject
Public gComboBox As OLEObject

Public gLastSearchValue As String

Public gHighlightBox As OLEObject
Public gHighlightCombo As OLEObject
Public gHighlightRange As Range

Public gMaskSheet As Worksheet
Public gMaskColumn As String
Public gMaskType As String
Public gLastMaskColumn As String


'=======================================================
' 1. SMART DATA MANAGER
'=======================================================

Sub ShowSmartDataManager(control As IRibbonControl)
    Dim rng As Range
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Ask user to select data range
    On Error Resume Next
    Set rng = Application.InputBox( _
    "Please select your Data Range including Headers." & Chr(13) & Chr(13) & _
    "INSTRUCTIONS:" & Chr(13) & _
    "1. Make sure the 1st Row is the HEADER." & Chr(13) & _
    "2. Include all columns you want to manage." & Chr(13) & _
    "3. Do not include merged cells.", _
    "Select Data with Header", _
    "", , , , , 8)
    On Error GoTo 0

    If rng Is Nothing Then Exit Sub

    Set g_DataRange = rng
    frmDataManager.InitForm rng
    frmDataManager.Show
End Sub

'=======================================================
' 2. SMART DATA EXPORTER
'=======================================================

Public Sub ShowSmartDataExporter(control As Object)

    Dim rng As Range
    Dim ws As Worksheet

    Set ws = ActiveSheet

    On Error Resume Next

    Set rng = Application.InputBox( _
        "Please select your Data Range including Headers." & Chr(13) & Chr(13) & _
        "INSTRUCTIONS:" & Chr(13) & _
        "1. Include HEADERS in the first row." & Chr(13) & _
        "2. Include the column you want to bulk export." & Chr(13) & _
        "3. Do not include merged cells.", _
        "Select Data for Bulk Export", _
        "", , , , , 8)

    On Error GoTo 0

    If rng Is Nothing Then Exit Sub

    Set g_DataRange = rng

    frmDataExporter.InitForm rng

    frmDataExporter.Show

End Sub



'=======================================================
' MODULE: modSmartDataManager
' SMART FILTER - XLBOOSTER STYLE
'=======================================================

'=======================================================
' SMART FILTER MAIN
'=======================================================

Sub ShowSmartFilter(control As IRibbonControl)

    Dim rng As Range
    Dim ws As Worksheet

    Set ws = ActiveSheet

    '-----------------------------------
    ' SELECT DATA
    '-----------------------------------
    On Error Resume Next


    Set rng = Application.InputBox( _
        "Please select your Data Range including Headers." & vbCrLf & vbCrLf & _
        "INSTRUCTIONS:" & vbCrLf & _
        "1. First row must contain HEADERS." & vbCrLf & _
        "2. Do not include merged cells." & vbCrLf & _
        "3. Select full table data.", _
        "Smart Filter - Select Data", _
        Type:=8)



    On Error GoTo 0

    If rng Is Nothing Then Exit Sub

    Set g_DataRange = rng

    '-----------------------------------
    ' CREATE DASHBOARD SHEET
    '-----------------------------------
    Dim dashWs As Worksheet

    Set dashWs = Sheets.Add(After:=Sheets(Sheets.count))

    dashWs.Name = "Dashboard_" & Format(Now, "HHMMSS")

    Set g_FilterSheet = dashWs

    BuildSmartFilterDashboard rng, dashWs

    ' start live watcher
    StartSmartFilterWatcher

    MsgBox _
        "Smart Filter Dashboard Ready!" & vbCrLf & vbCrLf & _
        "Type in search box to filter instantly.", _
        vbInformation, _
        "XLSmart"

End Sub


'=======================================================
' BUILD DASHBOARD
'=======================================================

Sub BuildSmartFilterDashboard(srcRange As Range, dashWs As Worksheet)

    Dim lastCol As Long
    Dim i As Long

    lastCol = srcRange.Columns.count

    Application.ScreenUpdating = False

    dashWs.Cells.Clear

    '===================================================
    ' TITLES
    '===================================================

    With dashWs.Range("B2")

        .Value = "1. Select Column"

        .Font.Bold = True

        .Font.Size = 11

    End With

    With dashWs.Range("F2")

        .Value = "2. Enter Search Text"

        .Font.Bold = True

        .Font.Size = 11

    End With


    '===================================================
    ' COMBOBOX
    '===================================================

    Dim cmb As OLEObject

    Set cmb = dashWs.OLEObjects.Add( _
                ClassType:="Forms.ComboBox.1", _
                left:=dashWs.Range("B3").left, _
                Top:=dashWs.Range("B3").Top, _
                Width:=170, _
                Height:=24)

    cmb.Name = "smartFilterCombo"

    Set gComboBox = cmb

    Dim cb As MSForms.ComboBox

    Set cb = cmb.Object

    For i = 1 To lastCol

        cb.AddItem srcRange.Cells(1, i).Value

    Next i

    cb.ListIndex = 0

    cb.Font.Size = 10


    '===================================================
    ' SEARCH BOX
    '===================================================

    Dim txt As OLEObject

    Set txt = dashWs.OLEObjects.Add( _
                ClassType:="Forms.TextBox.1", _
                left:=dashWs.Range("F3").left, _
                Top:=dashWs.Range("F3").Top, _
                Width:=180, _
                Height:=24)

    txt.Name = "smartSearchBox"

    txt.Object.Font.Size = 10

    Set gSearchBox = txt


    '===================================================
    ' EXPORT BUTTON
    '===================================================

    Dim shpBtn As Shape

    Set shpBtn = dashWs.Shapes.AddShape( _
                msoShapeRoundedRectangle, _
                dashWs.Range("J3").left, _
                dashWs.Range("J3").Top - 2, _
                80, _
                24)


    

    With shpBtn

    .Name = "smartExportBtn"

    .TextFrame.Characters.text = "Export as Excel"

    .Fill.ForeColor.RGB = RGB(44, 62, 80)

    .TextFrame.Characters.Font.Color = RGB(255, 255, 255)

    .TextFrame.Characters.Font.Bold = True

    .OnAction = "ExportSmartFilterData"

    End With






    '===================================================
    ' COPY DATA
    '===================================================

    dashWs.Range("B5").Resize( _
        srcRange.Rows.count, _
        srcRange.Columns.count).Value = srcRange.Value

    Dim dataRg As Range

    Set dataRg = dashWs.Range("B5").CurrentRegion

    Set g_DataRange = dataRg
    gLastSearchValue = ""


    '===================================================
    ' TABLE DESIGN
    '===================================================

    With dataRg

        .Font.Name = "Calibri"

        .Font.Size = 10

        .Borders.LineStyle = xlContinuous

    End With


    With dataRg.Rows(1)

    .Font.Bold = True

    .Interior.Color = RGB(36, 64, 98)

    .Font.Color = RGB(255, 255, 255)

    .HorizontalAlignment = xlCenter

    .VerticalAlignment = xlCenter

    End With


    dataRg.Rows(1).RowHeight = 24



    dataRg.Columns.AutoFit

    Application.ScreenUpdating = True

End Sub


'=======================================================
' LIVE FILTER ENGINE
'=======================================================

Public Sub StartSmartFilterWatcher()

    On Error Resume Next

    If gSearchBox Is Nothing Then Exit Sub

    Dim currentText As String

    currentText = Trim(gSearchBox.Object.text)

    If currentText <> gLastSearchValue Then

        gLastSearchValue = currentText

        RunSmartFilter

    End If

    Application.OnTime _
        Now + TimeValue("00:00:01"), _
        "StartSmartFilterWatcher"

End Sub


'=======================================================
' APPLY FILTER
'=======================================================

Public Sub RunSmartFilter()

    If g_DataRange Is Nothing Then Exit Sub

    Dim filterCol As String
    Dim searchVal As String

    filterCol = Trim(gComboBox.Object.Value)

    searchVal = Trim(gSearchBox.Object.text)

    Dim filterIndex As Long
    Dim i As Long

    filterIndex = 0

    '-----------------------------------
    ' FIND COLUMN
    '-----------------------------------
    For i = 1 To g_DataRange.Columns.count

        If Trim(g_DataRange.Cells(1, i).Value) = filterCol Then

            filterIndex = i

            Exit For

        End If

    Next i

    If filterIndex = 0 Then Exit Sub

    Application.ScreenUpdating = False

    On Error Resume Next

    ' clear old filter
    If g_DataRange.Parent.FilterMode Then

        g_DataRange.Parent.ShowAllData

    End If

    ' apply filter
    If searchVal <> "" Then

        g_DataRange.AutoFilter _
            Field:=filterIndex, _
            Criteria1:="*" & searchVal & "*"

    End If

    On Error GoTo 0

    Application.ScreenUpdating = True

End Sub


'=======================================================
' CLEAR FILTER
'=======================================================

Public Sub ClearSmartFilter()

    If g_DataRange Is Nothing Then Exit Sub

    On Error Resume Next

    If g_DataRange.Parent.FilterMode Then

        g_DataRange.Parent.ShowAllData

    End If

    gSearchBox.Object.text = ""

    gLastSearchValue = ""

    On Error GoTo 0

End Sub


'=======================================================
' EXPORT FILTERED DATA
'=======================================================

Public Sub ExportSmartFilterData()

    If g_DataRange Is Nothing Then Exit Sub

    Dim ans As VbMsgBoxResult

    ans = MsgBox( _
        "Do you want to export filtered data as separate Excel file?" & vbCrLf & vbCrLf & _
        "YES = Save XLSM file" & vbCrLf & _
        "NO = Create new worksheet", _
        vbYesNoCancel + vbQuestion, _
        "Export Filtered Data")

    If ans = vbCancel Then Exit Sub

    Dim exportRg As Range

    On Error Resume Next

    Set exportRg = g_DataRange.SpecialCells(xlCellTypeVisible)

    On Error GoTo 0

    If exportRg Is Nothing Then

        MsgBox "No filtered data found.", vbExclamation

        Exit Sub

    End If


    '===================================================
    ' EXPORT FILE
    '===================================================

    If ans = vbYes Then

        Dim fd As FileDialog
        Dim savePath As String

        Set fd = Application.FileDialog(msoFileDialogSaveAs)

        fd.Title = "Save Filtered File"

        fd.InitialFileName = _
            "FilteredData_" & Format(Now, "HHMMSS") & ".xlsm"

        If fd.Show <> -1 Then Exit Sub

        savePath = fd.SelectedItems(1)

        Dim newWb As Workbook
        Dim newWs As Worksheet

        Set newWb = Workbooks.Add

        Set newWs = newWb.Sheets(1)

        exportRg.Copy

        newWs.Range("A1").PasteSpecial xlPasteValues

        newWs.Range("A1").PasteSpecial xlPasteFormats

        newWs.Cells.Columns.AutoFit

        Application.DisplayAlerts = False

        newWb.SaveAs _
            Filename:=savePath, _
            FileFormat:=xlOpenXMLWorkbookMacroEnabled

        Application.DisplayAlerts = True

        MsgBox "Filtered file exported successfully.", vbInformation

    Else

        '===================================================
        ' EXPORT TO NEW SHEET
        '===================================================

        Dim ws As Worksheet

        Set ws = Sheets.Add(After:=Sheets(Sheets.count))

        ws.Name = "Filtered_" & Format(Now, "HHMMSS")

        exportRg.Copy

        ws.Range("A1").PasteSpecial xlPasteValues

        ws.Range("A1").PasteSpecial xlPasteFormats

        ws.Cells.Columns.AutoFit

        MsgBox "Filtered data inserted into new sheet.", vbInformation

    End If

End Sub

'=======================================================
' 4. SMART HIGHLIGHTER
'=======================================================

Sub ShowSmartHighlighter(control As IRibbonControl)

    Dim rng As Range
    Dim ws As Worksheet

    Set ws = ActiveSheet

    On Error Resume Next

    Set rng = Application.InputBox( _
        "Please select your Data Range including Headers." & vbCrLf & vbCrLf & _
        "First row must be HEADER.", _
        "Highlighter Engine - Select Data", _
        "", Type:=8)

    On Error GoTo 0

    If rng Is Nothing Then Exit Sub

    Set gHighlightRange = rng

    Dim ans As VbMsgBoxResult

    ans = MsgBox( _
        "Do you want to create the Smart Highlighter on a brand NEW sheet?" & vbCrLf & vbCrLf & _
        "• Click [YES] to automatically create it on a New Sheet." & vbCrLf & _
        "• Click [NO] to manually select a specific cell on an existing sheet.", _
        vbYesNoCancel + vbQuestion, _
        "Choose Destination - XLBooster")

    If ans = vbCancel Then Exit Sub

    Dim dashWs As Worksheet

    If ans = vbYes Then

        Set dashWs = Sheets.Add(After:=Sheets(Sheets.count))

        dashWs.Name = "Highlight_" & Format(Now, "HHMMSS")

    Else

        Set dashWs = ActiveSheet

    End If

    Application.ScreenUpdating = False

    dashWs.Cells.Clear

    '===================================================
    ' TITLES
    '===================================================

    With dashWs.Range("B2")

        .Value = "1. Select Column"

        .Font.Bold = True

        .Font.Size = 11

    End With

    With dashWs.Range("F2")

        .Value = "2. Enter Search Text"

        .Font.Bold = True

        .Font.Size = 11

    End With

    '===================================================
    ' COMBOBOX
    '===================================================

    Set gHighlightCombo = dashWs.OLEObjects.Add( _
        ClassType:="Forms.ComboBox.1", _
        left:=dashWs.Range("B3").left, _
        Top:=dashWs.Range("B3").Top, _
        Width:=170, _
        Height:=24)

    gHighlightCombo.Name = "cmbHighlight"

    Dim i As Long

    gHighlightCombo.Object.AddItem "All Columns"

    For i = 1 To rng.Columns.count

        gHighlightCombo.Object.AddItem rng.Cells(1, i).Value

    Next i

    gHighlightCombo.Object.ListIndex = 0

    '===================================================
    ' SEARCH BOX
    '===================================================
    Dim searchBox As OLEObject

    Set searchBox = dashWs.OLEObjects.Add( _
    ClassType:="Forms.TextBox.1", _
    left:=dashWs.Range("F3").left, _
    Top:=dashWs.Range("F3").Top, _
    Width:=200, _
    Height:=24)

    With searchBox

    .Name = "txtHighlightSearch"

    .Object.Font.Size = 10

    .Object.BorderStyle = 1

    .Object.text = ""

    End With


    '===================================================
    ' EXPORT BUTTON
    '===================================================

    Dim shpBtn As Shape

    Set shpBtn = dashWs.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        dashWs.Range("J3").left, _
        dashWs.Range("J3").Top - 2, _
        80, _
        24)

    With shpBtn

        .Name = "highlightExportBtn"

        .TextFrame.Characters.text = "Export as Excel"

        .Fill.ForeColor.RGB = RGB(44, 62, 80)

        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)

        .TextFrame.Characters.Font.Bold = True

        .OnAction = "ExportHighlightedData"

    End With

    '===================================================
    ' COPY DATA
    '===================================================

    dashWs.Range("B5").Resize( _
        rng.Rows.count, _
        rng.Columns.count).Value = rng.Value

    Dim dataRg As Range

    Set dataRg = dashWs.Range("B5").CurrentRegion

    Set g_DataRange = dataRg

    '===================================================
    ' DESIGN
    '===================================================

    With dataRg

        .Font.Name = "Calibri"

        .Font.Size = 10

        .Borders.LineStyle = xlContinuous

    End With

    With dataRg.Rows(1)

        .Font.Bold = True

        .Interior.Color = RGB(36, 64, 98)

        .Font.Color = RGB(255, 255, 255)

        .HorizontalAlignment = xlCenter

        .VerticalAlignment = xlCenter

    End With

    dataRg.Rows(1).RowHeight = 24

    dataRg.Columns.AutoFit

    Application.ScreenUpdating = True
    
    LiveHighlightSearch

    MsgBox "Smart Highlighter Dashboard Ready!", vbInformation

End Sub


'=======================================================
' RUN HIGHLIGHTER
'=======================================================

Sub RunSmartHighlighter()

    If g_DataRange Is Nothing Then Exit Sub

    Dim ws As Worksheet

    Set ws = g_DataRange.Worksheet

    Dim searchTxt As String
    Dim selectedCol As String

    searchTxt = LCase(Trim(ws.OLEObjects("txtHighlightSearch").Object.text))

    selectedCol = Trim(ws.OLEObjects("cmbHighlight").Object.Value)

    Dim rw As Long
    Dim cl As Long

    g_DataRange.Offset(1).Interior.ColorIndex = xlNone

    If searchTxt = "" Then Exit Sub

    '====================================
    ' ALL COLUMNS
    '====================================

    If selectedCol = "All Columns" Then

        For rw = 2 To g_DataRange.Rows.count

            For cl = 1 To g_DataRange.Columns.count

                If LCase(g_DataRange.Cells(rw, cl).Value) Like "*" & searchTxt & "*" Then

                    g_DataRange.Rows(rw).Interior.Color = RGB(255, 255, 120)

                    Exit For

                End If

            Next cl

        Next rw

    Else

        '================================
        ' SINGLE COLUMN
        '================================

        Dim filterCol As Long

        For cl = 1 To g_DataRange.Columns.count

            If Trim(g_DataRange.Cells(1, cl).Value) = selectedCol Then

                filterCol = cl

                Exit For

            End If

        Next cl

        If filterCol = 0 Then Exit Sub

        For rw = 2 To g_DataRange.Rows.count

            If LCase(g_DataRange.Cells(rw, filterCol).Value) Like "*" & searchTxt & "*" Then

                g_DataRange.Rows(rw).Interior.Color = RGB(255, 255, 120)

            End If

        Next rw

    End If

End Sub


'=======================================================
' EXPORT HIGHLIGHTED DATA
'=======================================================
Sub ExportHighlightedData()

    If g_DataRange Is Nothing Then Exit Sub

    Dim ans As VbMsgBoxResult

    ans = MsgBox( _
    "Do you want to save highlighted rows as a new Excel file?" & vbCrLf & vbCrLf & _
    "YES = Save as XLSM file" & vbCrLf & _
    "NO = Export to new sheet", _
    vbYesNoCancel + vbQuestion, _
    "Export Highlighted Data")

    If ans = vbCancel Then Exit Sub

    Dim ws As Worksheet
    Dim exportWs As Worksheet

    Set ws = g_DataRange.Worksheet

    Application.ScreenUpdating = False

    Set exportWs = Worksheets.Add(After:=Sheets(Sheets.count))

    exportWs.Name = "Highlighted_" & Format(Now, "HHMMSS")

    ' copy only highlighted rows
    g_DataRange.Rows(1).Copy exportWs.Range("A1")

    Dim lastRow As Long
    Dim rw As Long
    Dim nextRow As Long

    lastRow = g_DataRange.Rows.count

    nextRow = 2

    For rw = 2 To lastRow

        If g_DataRange.Rows(rw).Interior.Color = RGB(255, 255, 120) Then

            g_DataRange.Rows(rw).Copy _
            exportWs.Rows(nextRow)

            nextRow = nextRow + 1

        End If

    Next rw

    exportWs.Columns.AutoFit

    ' SAVE FILE
    If ans = vbYes Then

        Dim fPath As Variant

        fPath = Application.GetSaveAsFilename( _
        FileFilter:="Excel Macro Workbook (*.xlsm), *.xlsm")

        If fPath <> False Then

            exportWs.Parent.SaveAs fPath, 52

            MsgBox "Highlighted rows exported successfully!", vbInformation

        End If

    Else

        MsgBox "Highlighted rows exported to new sheet!", vbInformation

    End If

    Application.ScreenUpdating = True

End Sub


Sub LiveHighlightSearch()

    On Error Resume Next

    Dim currentTxt As String

    currentTxt = LCase(Trim( _
        g_DataRange.Worksheet.OLEObjects("txtHighlightSearch").Object.text))

    If currentTxt <> gLastSearchValue Then

        Application.ScreenUpdating = False

        gLastSearchValue = currentTxt

        RunSmartHighlighter

        Application.ScreenUpdating = True

    End If

    Application.OnTime _
        Now + TimeValue("00:00:01"), _
        "LiveHighlightSearch"

End Sub






















'=======================================================
' 5. SMART MASKING
'=======================================================

Sub ShowSmartMasking(control As IRibbonControl)

    Dim rng As Range
    Dim ws As Worksheet

    Set ws = ActiveSheet

    On Error Resume Next

    Set rng = Application.InputBox( _
        "Select your full data range including headers.", _
        "Smart Masking - Select Data", _
        Type:=8)

    On Error GoTo 0

    If rng Is Nothing Then Exit Sub

    SetupSmartMaskingDashboard rng

End Sub


'=======================================================
' BUILD MASK DASHBOARD
'=======================================================

Sub SetupSmartMaskingDashboard(srcRange As Range)

    Dim dashWs As Worksheet

    Set dashWs = Sheets.Add(After:=Sheets(Sheets.count))

    dashWs.Name = "Masking_" & Format(Now, "HHMMSS")

    Application.ScreenUpdating = False

    '-------------------------------
    ' TOP TITLES
    '-------------------------------

    dashWs.Range("B2").Value = "1. Select Column"
    dashWs.Range("F2").Value = "2. Select Mask Type"

    dashWs.Range("B2").Font.Bold = True
    dashWs.Range("F2").Font.Bold = True

    '-------------------------------
    ' COLUMN DROPDOWN
    '-------------------------------

    Dim cmbCol As OLEObject

    Set cmbCol = dashWs.OLEObjects.Add( _
        ClassType:="Forms.ComboBox.1", _
        left:=dashWs.Range("B3").left, _
        Top:=dashWs.Range("B3").Top, _
        Width:=170, _
        Height:=24)

    cmbCol.Name = "cmbMaskColumn"

    Dim i As Long

    cmbCol.Object.AddItem "All Columns"

    For i = 1 To srcRange.Columns.count

        cmbCol.Object.AddItem srcRange.Cells(1, i).Value

    Next i

    cmbCol.Object.ListIndex = 0

    '-------------------------------
    ' MASK TYPE DROPDOWN
    '-------------------------------

    Dim cmbMask As OLEObject

    Set cmbMask = dashWs.OLEObjects.Add( _
        ClassType:="Forms.ComboBox.1", _
        left:=dashWs.Range("F3").left, _
        Top:=dashWs.Range("F3").Top, _
        Width:=220, _
        Height:=24)

    cmbMask.Name = "cmbMaskType"

    With cmbMask.Object

        .AddItem "0. None (Original Data)"
        .AddItem "1. Full Mask (********)"
        .AddItem "2. Name Partial (s***** k********)"
        .AddItem "3. Mobile Number (+8801******89)"
        .AddItem "4. Email ID (sh***@gmail.com)"
        .AddItem "5. National ID Card (XXXX-XXXX-1234)"
        .AddItem "6. Bank Card (**** **** **** 4589)"
        .AddItem "7. Bank Account (XXXXXX1234)"
        .AddItem "8. Passport Number (EH****567)"
        .AddItem "9. Date of Birth (XX-XX-YYYY)"

        .ListIndex = 0

    End With


'-------------------------------
' APPLY BUTTON
'-------------------------------
    Dim applyBtn As Shape

    Set applyBtn = dashWs.Shapes.AddShape( _
    msoShapeRoundedRectangle, _
    dashWs.Range("J3").left, _
    dashWs.Range("J3").Top, _
    65, _
    24)

    With applyBtn

    .Name = "btnApplyMask"

    .TextFrame.Characters.text = "Apply Mask"

    .Fill.ForeColor.RGB = RGB(52, 152, 219)

    .TextFrame.Characters.Font.Color = RGB(255, 255, 255)

    .TextFrame.Characters.Font.Bold = True

    .OnAction = "RunSmartMasking"

    End With


'-------------------------------
' EXPORT BUTTON
'-------------------------------

    Dim shpBtn As Shape

    Set shpBtn = dashWs.Shapes.AddShape( _
    msoShapeRoundedRectangle, _
    dashWs.Range("M3").left, _
    dashWs.Range("M3").Top, _
    70, _
    24)

    With shpBtn

    .Name = "btnExportMasked"

    .TextFrame.Characters.text = "Export Masked Data"

    .Fill.ForeColor.RGB = RGB(44, 62, 80)

    .TextFrame.Characters.Font.Color = RGB(255, 255, 255)

    .TextFrame.Characters.Font.Bold = True

    .OnAction = "ExportMaskedDashboard"

    End With

    '-------------------------------
    ' COPY DATA
    '-------------------------------

    dashWs.Range("B5").Resize( _
        srcRange.Rows.count, _
        srcRange.Columns.count).Value = srcRange.Value

    Dim dataRg As Range

    Set dataRg = dashWs.Range("B5").CurrentRegion

    Set g_DataRange = dataRg

    Set gMaskSheet = dashWs
    gMaskType = ""
    gMaskColumn = ""
    gLastMaskColumn = ""

    '-------------------------------
    ' TABLE STYLE
    '-------------------------------

    With dataRg

        .Font.Name = "Calibri"

        .Font.Size = 10

        .Borders.LineStyle = xlContinuous

    End With

    With dataRg.Rows(1)

        .Font.Bold = True

        .Interior.Color = RGB(36, 64, 98)

        .Font.Color = RGB(255, 255, 255)

        .HorizontalAlignment = xlCenter

        .VerticalAlignment = xlCenter

    End With

    dataRg.Rows(1).RowHeight = 24

    dataRg.Columns.AutoFit

    Application.ScreenUpdating = True

    

    MsgBox "Smart Masking Dashboard Ready!", vbInformation

End Sub





'=======================================================
' RUN MASKING
'=======================================================

Sub RunSmartMasking()

    If g_DataRange Is Nothing Then Exit Sub

    Dim ws As Worksheet

    Set ws = g_DataRange.Worksheet

    Dim selectedCol As String
    Dim maskType As String

    selectedCol = ws.OLEObjects("cmbMaskColumn").Object.Value

    maskType = ws.OLEObjects("cmbMaskType").Object.Value

    If maskType = "0. None (Original Data)" Then Exit Sub

    Dim rw As Long
    Dim cl As Long

    For rw = 2 To g_DataRange.Rows.count

        For cl = 1 To g_DataRange.Columns.count

            Dim hdr As String

            hdr = Trim(g_DataRange.Cells(1, cl).Value)

            If selectedCol = "All Columns" Or hdr = selectedCol Then

                g_DataRange.Cells(rw, cl).Value = _
                    GetMaskedValue( _
                    CStr(g_DataRange.Cells(rw, cl).Value), _
                    maskType)

            End If

        Next cl

    Next rw

End Sub


'=======================================================
' GET MASKED VALUE
'=======================================================

Function GetMaskedValue(txt As String, maskType As String) As String

    Dim result As String

    result = txt

    Select Case True

        '-----------------------------------
        ' FULL MASK
        '-----------------------------------
        Case InStr(maskType, "Full Mask") > 0

            result = String(Len(txt), "*")


        '-----------------------------------
        ' NAME MASK
        ' sharif khandakar
        ' s***** k********
        '-----------------------------------
        Case InStr(maskType, "Name Partial") > 0

            Dim parts() As String
            Dim finalName As String
            Dim i As Long

            parts = Split(txt, " ")

            finalName = ""

            For i = LBound(parts) To UBound(parts)

                If Len(parts(i)) > 1 Then

                    finalName = finalName & _
                        left(parts(i), 1) & _
                        String(Len(parts(i)) - 1, "*") & " "

                Else

                    finalName = finalName & parts(i) & " "

                End If

            Next i

            result = Trim(finalName)


        '-----------------------------------
        ' MOBILE NUMBER
        ' +8801******89
        '-----------------------------------
        Case InStr(maskType, "Mobile") > 0

            If Len(txt) >= 11 Then

                result = left(txt, 5) & _
                        String(Len(txt) - 7, "*") & _
                        right(txt, 2)

            End If


        '-----------------------------------
        ' EMAIL MASK
        '-----------------------------------
        Case InStr(maskType, "Email") > 0

            Dim atPos As Long

            atPos = InStr(txt, "@")

            If atPos > 3 Then

                result = left(txt, 2) & "***" & Mid(txt, atPos)

            End If


        '-----------------------------------
        ' NATIONAL ID CARD
        '-----------------------------------
        Case InStr(maskType, "National") > 0

            If Len(txt) >= 4 Then

                result = "XXXX-XXXX-" & right(txt, 4)

            End If


        '-----------------------------------
        ' BANK CARD
        '-----------------------------------
        Case InStr(maskType, "Bank Card") > 0

            If Len(txt) >= 4 Then

                result = "**** **** **** " & right(txt, 4)

            End If


        '-----------------------------------
        ' BANK ACCOUNT
        '-----------------------------------
        Case InStr(maskType, "Bank Account") > 0

            If Len(txt) >= 4 Then

                result = "XXXXXX" & right(txt, 4)

            End If


        '-----------------------------------
        ' PASSPORT
        ' EH****567
        '-----------------------------------
        Case InStr(maskType, "Passport") > 0

            If Len(txt) >= 5 Then

                result = left(txt, 2) & _
                        String(Len(txt) - 5, "*") & _
                        right(txt, 3)

            End If


        '-----------------------------------
        ' DATE OF BIRTH
        '-----------------------------------
        Case InStr(maskType, "Date") > 0

            result = "XX-XX-YYYY"

    End Select

    GetMaskedValue = result

End Function


'=======================================================
' EXPORT MASKED DATA
'=======================================================

Sub ExportMaskedDashboard()

    If g_DataRange Is Nothing Then Exit Sub

    Dim ans As VbMsgBoxResult

    ans = MsgBox( _
        "Do you want to export masked data as PDF file?" & vbCrLf & vbCrLf & _
        "YES = Export as PDF" & vbCrLf & _
        "NO = Export to new Sheet", _
        vbYesNoCancel + vbQuestion)

    If ans = vbCancel Then Exit Sub

    Dim ws As Worksheet
    Dim exportWs As Worksheet

    Set ws = gMaskSheet

    Application.ScreenUpdating = False

    '-----------------------------------
    ' CREATE CLEAN EXPORT SHEET
    '-----------------------------------

    Set exportWs = Sheets.Add(After:=Sheets(Sheets.count))

    exportWs.Name = "MaskedExport_" & Format(Now, "HHMMSS")

    ' copy only table
    g_DataRange.Copy exportWs.Range("A1")

    exportWs.Columns.AutoFit

    '-----------------------------------
    ' PDF EXPORT
    '-----------------------------------

    If ans = vbYes Then

        Dim fPath As Variant

        fPath = Application.GetSaveAsFilename( _
            FileFilter:="PDF Files (*.pdf), *.pdf")

        If fPath <> False Then

    With exportWs.PageSetup

        .Orientation = xlLandscape

        .Zoom = False

        .FitToPagesWide = 1

        .FitToPagesTall = False
        .PrintTitleRows = "$1:$1"

        .LeftMargin = Application.InchesToPoints(0.3)

        .RightMargin = Application.InchesToPoints(0.3)

        .TopMargin = Application.InchesToPoints(0.5)

        .BottomMargin = Application.InchesToPoints(0.5)

        .CenterHorizontally = True

    End With

        exportWs.Columns.AutoFit

        exportWs.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=fPath, _
        Quality:=xlQualityStandard

            MsgBox "PDF Export Complete!", vbInformation

        End If

    Else

        MsgBox "Masked data exported to new sheet.", vbInformation

    End If

    Application.ScreenUpdating = True

End Sub

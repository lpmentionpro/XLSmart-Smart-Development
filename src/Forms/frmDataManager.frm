VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataManager 
   Caption         =   "Smart Data Manager"
   ClientHeight    =   13440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21990
   OleObjectBlob   =   "frmDataManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDataManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================
' USERFORM : frmDataManager
' Smart Data Manager - Full Updated Version
'========================================================

Option Explicit
Dim isDarkMode As Boolean
Private m_DataRange As Range
Private m_Headers() As String
Private m_ColCount As Long

Private m_OrigData As Variant
Private m_FiltData As Variant

Private m_SelectedRow As Long
Private m_DarkMode As Boolean

'========================================================
' INIT FORM
'========================================================
Public Sub InitForm(srcRange As Range)

    Set m_DataRange = srcRange

    m_ColCount = srcRange.Columns.count

    ReDim m_Headers(1 To m_ColCount)

    Dim i As Long

    For i = 1 To m_ColCount

        m_Headers(i) = Trim(CStr(srcRange.Cells(1, i).Value))

    Next i

    LoadSheetData

    m_SelectedRow = 0
    m_DarkMode = False

    BuildDynamicFields
    PopulateAllCombos
    RenderDataList

End Sub

'========================================================
' USERFORM LOAD
'========================================================
Private Sub UserForm_Initialize()

    lstData.MultiSelect = fmMultiSelectSingle
    lstData.ListStyle = fmListStylePlain
    lstData.IntegralHeight = False

    panelFields.ScrollBars = fmScrollBarsVertical

End Sub

'========================================================
' LOAD SHEET DATA
'========================================================
Private Sub LoadSheetData()

    Dim totalRows As Long

    totalRows = m_DataRange.Rows.count - 1

    If totalRows <= 0 Then Exit Sub

    m_OrigData = m_DataRange.Offset(1).Resize(totalRows).Value

    m_FiltData = m_OrigData

End Sub

'========================================================
' BUILD DYNAMIC FIELDS
'========================================================
Private Sub BuildDynamicFields()

    Dim ctrl As control
    Dim arrTmp() As String

    Dim cnt As Long
    cnt = 0

    For Each ctrl In panelFields.Controls

        cnt = cnt + 1

        ReDim Preserve arrTmp(1 To cnt)

        arrTmp(cnt) = ctrl.Name

    Next ctrl

    Dim i As Long

    For i = 1 To cnt

        panelFields.Controls.Remove arrTmp(i)

    Next i

    Dim topPos As Single

    topPos = 8

    Dim lbl As MSForms.Label
    Dim txt As MSForms.TextBox
    Dim cbo As MSForms.ComboBox

    For i = 1 To m_ColCount

        Set lbl = panelFields.Controls.Add("Forms.Label.1", "lbl_" & i)

        With lbl

            .Caption = m_Headers(i)

            .left = 8
            .Top = topPos + 3

            .Width = 100
            .Height = 18

            .Font.Bold = True
            .Font.Size = 9

        End With

        If IsDropdownColumn(i) Then

            Set cbo = panelFields.Controls.Add("Forms.ComboBox.1", "fld_" & i)

            With cbo

                .left = 115
                .Top = topPos

                .Width = 190
                .Height = 20

                .Style = fmStyleDropDownCombo

                .Font.Size = 9

            End With

            LoadUniqueValues cbo, i

        Else

            Set txt = panelFields.Controls.Add("Forms.TextBox.1", "fld_" & i)

            With txt

                .left = 115
                .Top = topPos

                .Width = 190
                .Height = 20

                .Font.Size = 9

            End With

        End If

        topPos = topPos + 28

    Next i

    panelFields.ScrollBars = fmScrollBarsVertical
    panelFields.ScrollHeight = topPos + 30

End Sub

'========================================================
' SMART DROPDOWN DETECT
'========================================================
Private Function IsDropdownColumn(colIdx As Long) As Boolean

    Dim h As String

    h = LCase(m_Headers(colIdx))

    If InStr(h, "status") > 0 Then IsDropdownColumn = True
    If InStr(h, "category") > 0 Then IsDropdownColumn = True
    If InStr(h, "city") > 0 Then IsDropdownColumn = True
    If InStr(h, "state") > 0 Then IsDropdownColumn = True
    If InStr(h, "region") > 0 Then IsDropdownColumn = True
    If InStr(h, "type") > 0 Then IsDropdownColumn = True

End Function

'========================================================
' LOAD UNIQUE VALUES
'========================================================
Private Sub LoadUniqueValues(cbo As MSForms.ComboBox, colIdx As Long)

    Dim dict As Object

    Set dict = CreateObject("Scripting.Dictionary")

    Dim r As Long
    Dim v As String

    For r = 1 To UBound(m_OrigData, 1)

        v = Trim(CStr(m_OrigData(r, colIdx)))

        If v <> "" Then

            If Not dict.exists(v) Then
                dict.Add v, 1
            End If

        End If

    Next r

    Dim k

    For Each k In dict.Keys

        cbo.AddItem k

    Next k

End Sub

'========================================================
' POPULATE COMBOS
'========================================================
Private Sub PopulateAllCombos()

    Dim i As Long

    cboFilterCol.Clear
    cboSortCol.Clear
    cboTopValCol.Clear
    cboGTLTCol.Clear

    cboFilterCol.AddItem "Select Column..."
    cboGTLTCol.AddItem "Column..."

    For i = 1 To m_ColCount

        cboFilterCol.AddItem m_Headers(i)

        cboSortCol.AddItem m_Headers(i)

        cboTopValCol.AddItem m_Headers(i)

        cboGTLTCol.AddItem m_Headers(i)

    Next i

    cboFilterCol.ListIndex = 0
    cboGTLTCol.ListIndex = 0

    If cboSortCol.ListCount > 0 Then
        cboSortCol.ListIndex = 0
    End If

    If cboTopValCol.ListCount > 0 Then
        cboTopValCol.ListIndex = 0
    End If

    cboSortOrder.Clear

    cboSortOrder.AddItem "Ascending"
    cboSortOrder.AddItem "Descending"

    cboSortOrder.ListIndex = 0

    cboTopN.Clear

    cboTopN.AddItem "Top"
    cboTopN.AddItem "Bottom"

    cboTopN.ListIndex = 0



    cboGTLTOp.Clear

    cboGTLTOp.AddItem ">"
    cboGTLTOp.AddItem "<"
    cboGTLTOp.AddItem "="
    cboGTLTOp.AddItem ">="
    cboGTLTOp.AddItem "<="

    cboGTLTOp.ListIndex = 0

    txtTopCount.Value = "5"

End Sub

'========================================================
' RENDER DATA
'========================================================
Private Sub RenderDataList()

    lstData.Clear

    If IsArrayEmpty(m_FiltData) Then Exit Sub

    Dim colWidthStr As String
    Dim i As Long

    colWidthStr = ""

    For i = 1 To m_ColCount

        colWidthStr = colWidthStr & "70 pt;"

    Next i

    lstData.ColumnCount = m_ColCount

    lstData.ColumnWidths = colWidthStr

    lstData.ListStyle = fmListStylePlain
    lstData.MultiSelect = fmMultiSelectSingle

    lstData.Font.Name = "Consolas"
    lstData.Font.Size = 9

    lstData.IntegralHeight = False

    lstData.List = m_FiltData

    lblRowCount.Caption = _
    "Total Records: " & lstData.ListCount

End Sub

'========================================================
' EMPTY ARRAY CHECK
'========================================================
Private Function IsArrayEmpty(arr As Variant) As Boolean

    On Error GoTo errHandle

    If IsEmpty(arr) Then
        IsArrayEmpty = True
        Exit Function
    End If

    Dim x As Long

    x = UBound(arr, 1)

    IsArrayEmpty = False

    Exit Function

errHandle:

    IsArrayEmpty = True

End Function

'========================================================
' LIST CLICK
'========================================================
Private Sub lstData_Click()

    If lstData.ListIndex < 0 Then Exit Sub

    m_SelectedRow = lstData.ListIndex + 1

    Dim i As Long
    Dim ctrl As control

    For i = 1 To m_ColCount

        Set ctrl = panelFields.Controls("fld_" & i)

        ctrl.Value = lstData.List(lstData.ListIndex, i - 1)

    Next i

End Sub

'========================================================
' LIVE SEARCH
'========================================================
Private Sub txtSearch_Change()

    Dim s As String

    s = LCase(Trim(txtSearch.text))

    If s = "" Then

        m_FiltData = m_OrigData

        RenderDataList

        Exit Sub

    End If

    Dim tempArr() As Variant

    ReDim tempArr(1 To UBound(m_OrigData, 1), 1 To m_ColCount)

    Dim r As Long
    Dim c As Long
    Dim cnt As Long

    cnt = 0

    For r = 1 To UBound(m_OrigData, 1)

        Dim found As Boolean

        found = False

        For c = 1 To m_ColCount

            If InStr(1, _
            LCase(CStr(m_OrigData(r, c))), s) > 0 Then

                found = True
                Exit For

            End If

        Next c

        If found Then

            cnt = cnt + 1

            For c = 1 To m_ColCount

                tempArr(cnt, c) = m_OrigData(r, c)

            Next c

        End If

    Next r

    If cnt = 0 Then

        lstData.Clear

        lblRowCount.Caption = "Total Records: 0"

        Exit Sub

    End If

    ReDim m_FiltData(1 To cnt, 1 To m_ColCount)

    For r = 1 To cnt

        For c = 1 To m_ColCount

            m_FiltData(r, c) = tempArr(r, c)

        Next c

    Next r

    RenderDataList

End Sub

'========================================================
' FILTER COLUMN
'========================================================
Private Sub cboFilterCol_Change()

    cboFilterVal.Clear

    If cboFilterCol.ListIndex <= 0 Then Exit Sub

    Dim colIdx As Long

    colIdx = cboFilterCol.ListIndex

    Dim dict As Object

    Set dict = CreateObject("Scripting.Dictionary")

    Dim r As Long
    Dim v As String

    For r = 1 To UBound(m_OrigData, 1)

        v = Trim(CStr(m_OrigData(r, colIdx)))

        If v <> "" Then

            If Not dict.exists(v) Then

                dict.Add v, 1

            End If

        End If

    Next r

    Dim k

    For Each k In dict.Keys

        cboFilterVal.AddItem k

    Next k

    If cboFilterVal.ListCount > 0 Then
        cboFilterVal.ListIndex = 0
    End If

End Sub

'========================================================
' FILTER VALUE
'========================================================
Private Sub cboFilterVal_Change()

    If cboFilterCol.ListIndex <= 0 Then Exit Sub

    If cboFilterVal.Value = "" Then Exit Sub

    Dim colIdx As Long

    colIdx = cboFilterCol.ListIndex

    Dim tempArr() As Variant

    ReDim tempArr(1 To UBound(m_OrigData, 1), 1 To m_ColCount)

    Dim r As Long
    Dim c As Long
    Dim cnt As Long

    cnt = 0

    For r = 1 To UBound(m_OrigData, 1)

        If CStr(m_OrigData(r, colIdx)) = cboFilterVal.Value Then

            cnt = cnt + 1

            For c = 1 To m_ColCount

                tempArr(cnt, c) = m_OrigData(r, c)

            Next c

        End If

    Next r

    If cnt = 0 Then Exit Sub

    ReDim m_FiltData(1 To cnt, 1 To m_ColCount)

    For r = 1 To cnt

        For c = 1 To m_ColCount

            m_FiltData(r, c) = tempArr(r, c)

        Next c

    Next r

    RenderDataList

End Sub
'========================================================
' DATE FILTER
'========================================================
Private Sub ApplyDateFilter()

    If Trim(txtDateFrom.Value) = "" Then Exit Sub
    If Trim(txtDateTo.Value) = "" Then Exit Sub

    If Not IsDate(txtDateFrom.Value) Then

        MsgBox "Invalid FROM date."
        Exit Sub

    End If

    If Not IsDate(txtDateTo.Value) Then

        MsgBox "Invalid TO date."
        Exit Sub

    End If

    Dim fromDt As Date
    Dim toDt As Date

    fromDt = CDate(txtDateFrom.Value)
    toDt = CDate(txtDateTo.Value)

    Dim dateCol As Long
    Dim i As Long

    ' detect date column
    For i = 1 To m_ColCount

        If InStr(1, LCase(m_Headers(i)), "date") > 0 Then

            dateCol = i
            Exit For

        End If

    Next i

    If dateCol = 0 Then

        MsgBox "Date column not found."
        Exit Sub

    End If

    Dim tempArr() As Variant

    ReDim tempArr(1 To UBound(m_OrigData, 1), 1 To m_ColCount)

    Dim r As Long
    Dim c As Long
    Dim cnt As Long

    cnt = 0

    Dim rowDate As Date

    For r = 1 To UBound(m_OrigData, 1)

        If IsDate(m_OrigData(r, dateCol)) Then

            rowDate = CDate(m_OrigData(r, dateCol))

            If rowDate >= fromDt _
            And rowDate <= toDt Then

                cnt = cnt + 1

                For c = 1 To m_ColCount

                    tempArr(cnt, c) = m_OrigData(r, c)

                Next c

            End If

        End If

    Next r

    If cnt = 0 Then

        lstData.Clear

        lblRowCount.Caption = "Total Records: 0"

        MsgBox "No records found."

        Exit Sub

    End If

    ReDim m_FiltData(1 To cnt, 1 To m_ColCount)

    For r = 1 To cnt

        For c = 1 To m_ColCount

            m_FiltData(r, c) = tempArr(r, c)

        Next c

    Next r

    RenderDataList

End Sub

'========================================================
' APPLY DATE FILTER BUTTON
'========================================================
Private Sub btnApplyDate_Click()

    ApplyDateFilter

End Sub

'========================================================
' SORT
'========================================================
Private Sub btnSort_Click()

    Dim colIdx As Long

    colIdx = cboSortCol.ListIndex + 1

    Dim sortOrder As XlSortOrder

    If cboSortOrder.Value = "Descending" Then
        sortOrder = xlDescending
    Else
        sortOrder = xlAscending
    End If

    m_DataRange.Sort _
    Key1:=m_DataRange.Columns(colIdx), _
    Order1:=sortOrder, _
    Header:=xlYes

    LoadSheetData
    RenderDataList

End Sub

'========================================================
' RESET
'========================================================
Private Sub btnResetSort_Click()

    txtSearch.text = ""

    m_FiltData = m_OrigData

    RenderDataList

End Sub

'========================================================
' CLEAR FORM
'========================================================
Private Sub btnClearForm_Click()

    Dim ctrl As control

    For Each ctrl In panelFields.Controls

        If TypeName(ctrl) = "TextBox" _
        Or TypeName(ctrl) = "ComboBox" Then

            ctrl.Value = ""

        End If

    Next ctrl

End Sub

'========================================================
' ADD
'========================================================
Private Sub btnAdd_Click()

    Dim nextRow As Long

    nextRow = m_DataRange.Rows.count + m_DataRange.Row

    Dim i As Long
    Dim ctrl As control

    For i = 1 To m_ColCount

        Set ctrl = panelFields.Controls("fld_" & i)

        Cells(nextRow, _
        m_DataRange.Column + i - 1).Value = ctrl.Value

    Next i

    Set m_DataRange = _
    m_DataRange.Resize(m_DataRange.Rows.count + 1)

    LoadSheetData
    RenderDataList

    MsgBox "New data added."

End Sub

'========================================================
' UPDATE
'========================================================
Private Sub btnUpdate_Click()

    If m_SelectedRow = 0 Then

        MsgBox "Select row first."
        Exit Sub

    End If

    Dim i As Long
    Dim ctrl As control

    For i = 1 To m_ColCount

        Set ctrl = panelFields.Controls("fld_" & i)

        m_DataRange.Cells(m_SelectedRow + 1, i).Value = ctrl.Value

    Next i

    LoadSheetData
    RenderDataList

    MsgBox "Updated."

End Sub

'========================================================
' DELETE
'========================================================
Private Sub btnDelete_Click()

    If m_SelectedRow = 0 Then

        MsgBox "Select row first."
        Exit Sub

    End If

    m_DataRange.Rows(m_SelectedRow + 1).Delete

    Set m_DataRange = _
    m_DataRange.Resize(m_DataRange.Rows.count - 1)

    LoadSheetData
    RenderDataList

    btnClearForm_Click

End Sub

'========================================================
' EXPORT EXCEL WITH FORMAT
'========================================================
Private Sub btnExportExcel_Click()

    Dim srcSheet As Worksheet
    Dim newBook As Workbook
    Dim newSheet As Worksheet

    Set srcSheet = ActiveSheet

    Application.ScreenUpdating = False

    ' create hidden workbook
    Set newBook = Workbooks.Add

    Set newSheet = newBook.Sheets(1)

    ' copy entire format
    srcSheet.Cells.Copy

    newSheet.Range("A1").PasteSpecial xlPasteFormats
    newSheet.Range("A1").PasteSpecial xlPasteColumnWidths

    Dim r As Long
    Dim c As Long

    ' headers
    For c = 1 To m_ColCount

        newSheet.Cells(1, c).Value = m_Headers(c)

    Next c

    ' clear old rows
    newSheet.Rows("2:" & newSheet.Rows.count).ClearContents

    ' insert filtered data
    For r = 1 To UBound(m_FiltData, 1)

        For c = 1 To m_ColCount

            newSheet.Cells(r + 1, c).Value = _
            m_FiltData(r, c)

        Next c

    Next r

    Dim lastRow As Long

    lastRow = _
    newSheet.Cells( _
    newSheet.Rows.count, 1).End(xlUp).Row

    ' borders
    With newSheet.Range( _
    newSheet.Cells(1, 1), _
    newSheet.Cells(lastRow, m_ColCount)).Borders

        .LineStyle = xlContinuous

    End With

    newSheet.Cells.EntireColumn.AutoFit

    Dim fd As FileDialog
    Dim savePath As String

    Set fd = _
    Application.FileDialog(msoFileDialogSaveAs)

    fd.Title = "Save Excel File"

    fd.InitialFileName = _
    ThisWorkbook.Path & "\Report.xlsx"

    If fd.Show <> -1 Then

        newBook.Close False
        Exit Sub

    End If

    savePath = fd.SelectedItems(1)

    If LCase(right(savePath, 5)) <> ".xlsx" Then

        savePath = savePath & ".xlsx"

    End If

    Application.DisplayAlerts = False

    newBook.SaveAs _
    Filename:=savePath, _
    FileFormat:=xlOpenXMLWorkbook

    newBook.Close False

    Application.DisplayAlerts = True

    Application.ScreenUpdating = True

    MsgBox "Excel exported successfully."

End Sub

'========================================================
' EXPORT PDF WITH ORIGINAL FORMAT
'========================================================
Private Sub btnExportPDF_Click()

    On Error GoTo ErrHandler

    Dim srcSheet As Worksheet
    Dim tmpSheet As Worksheet

    Dim lastRow As Long
    Dim lastCol As Long

    Dim r As Long
    Dim c As Long

    Dim fd As FileDialog
    Dim savePath As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set srcSheet = ActiveSheet

    ' copy original sheet with formatting
    srcSheet.Copy After:=Sheets(Sheets.count)

    Set tmpSheet = ActiveSheet

    tmpSheet.Name = _
    "PDF_" & Format(Now, "hhmmss")

    lastCol = m_ColCount

    lastRow = _
    tmpSheet.Cells( _
    tmpSheet.Rows.count, 1).End(xlUp).Row

    ' remove old data rows but keep header
    If lastRow >= 2 Then

        tmpSheet.Range( _
        tmpSheet.Rows(2), _
        tmpSheet.Rows(lastRow)).EntireRow.Delete

    End If

    ' insert filtered data
    For r = 1 To UBound(m_FiltData, 1)

        For c = 1 To m_ColCount

            tmpSheet.Cells(r + 1, c).Value = _
            m_FiltData(r, c)

        Next c

    Next r

    lastRow = _
    tmpSheet.Cells( _
    tmpSheet.Rows.count, 1).End(xlUp).Row

    ' borders
    With tmpSheet.Range( _
    tmpSheet.Cells(1, 1), _
    tmpSheet.Cells(lastRow, lastCol)).Borders

        .LineStyle = xlContinuous
        .Weight = xlThin

    End With

    ' alignment
    With tmpSheet.Range( _
    tmpSheet.Cells(1, 1), _
    tmpSheet.Cells(lastRow, lastCol))

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    ' header style
    With tmpSheet.Rows(1)

        .Font.Bold = True

        .HorizontalAlignment = xlCenter

    End With

    ' autofit
    tmpSheet.Cells.EntireColumn.AutoFit

    ' page setup
    With tmpSheet.PageSetup

        .Orientation = xlLandscape

        .Zoom = False

        .FitToPagesWide = 1
        .FitToPagesTall = False

        .PrintTitleRows = "$1:$1"

        .CenterHorizontally = True

        .LeftMargin = _
        Application.InchesToPoints(0.3)

        .RightMargin = _
        Application.InchesToPoints(0.3)

        .TopMargin = _
        Application.InchesToPoints(0.5)

        .BottomMargin = _
        Application.InchesToPoints(0.5)

    End With

    ' save dialog
    Set fd = _
    Application.FileDialog(msoFileDialogSaveAs)

    fd.Title = "Save PDF"

    fd.InitialFileName = _
    ThisWorkbook.Path & "\Report.pdf"

    If fd.Show <> -1 Then

        tmpSheet.Delete

        GoTo SafeExit

    End If

    savePath = fd.SelectedItems(1)

    If LCase(right(savePath, 4)) <> ".pdf" Then

        savePath = savePath & ".pdf"

    End If

    ' export pdf
    tmpSheet.ExportAsFixedFormat _
    Type:=xlTypePDF, _
    Filename:=savePath, _
    Quality:=xlQualityStandard

    tmpSheet.Delete

    MsgBox "PDF exported successfully."

SafeExit:

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Exit Sub

ErrHandler:

    MsgBox Err.Description

    Resume SafeExit

End Sub

'========================================================
' INSERT DATA
'========================================================
Private Sub btnInsertData_Click()

    Dim ans As VbMsgBoxResult

    ans = MsgBox( _
    "Do you want to insert preview data into a NEW worksheet?" & vbCrLf & vbCrLf & _
    "YES - Insert into New Sheet (Auto-Named)" & vbCrLf & _
    "NO - Insert into selected cell on Existing Sheet", _
    vbYesNoCancel + vbQuestion, _
    "Insert Preview Data")

    If ans = vbCancel Then Exit Sub

    If ans = vbYes Then

        Call ExportToNewSheet

    ElseIf ans = vbNo Then

        Call InsertIntoExistingSheet

    End If

End Sub

'========================================================
' EXPORT TO NEW SHEET
'========================================================
Private Sub ExportToNewSheet()

    Dim ws As Worksheet
    Dim shtName As String

    Dim usingFilter As Boolean
    Dim usingTop As Boolean

    usingFilter = False
    usingTop = False

    '--------------------------------
    ' detect active operation
    '--------------------------------
    If cboFilterCol.Value <> "" And _
       cboFilterVal.Value <> "" Then

        usingFilter = True

    End If

    If cboTopN.Value <> "" And _
       txtTopCount.Value <> "" And _
       cboTopValCol.Value <> "" Then

        usingTop = True

    End If

    '--------------------------------
    ' SHEET NAME
    '--------------------------------
    If usingFilter Then

        shtName = _
        cboFilterCol.Value & "_" & _
        cboFilterVal.Value

    ElseIf usingTop Then

        shtName = _
        cboTopN.Value & _
        txtTopCount.Value & "_" & _
        cboTopValCol.Value

    Else

        shtName = "FilteredData"

    End If

    '--------------------------------
    ' remove invalid chars
    '--------------------------------
    shtName = Replace(shtName, "/", "")
    shtName = Replace(shtName, "\", "")
    shtName = Replace(shtName, ":", "")
    shtName = Replace(shtName, "*", "")
    shtName = Replace(shtName, "?", "")
    shtName = Replace(shtName, "[", "")
    shtName = Replace(shtName, "]", "")

    ' max 31 chars
    shtName = left(shtName, 31)

    '--------------------------------
    ' create sheet AFTER last sheet
    '--------------------------------
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.count))

    On Error Resume Next

    ws.Name = shtName

    If Err.Number <> 0 Then

        ws.Name = shtName & "_" & Format(Now, "hhmmss")
        Err.Clear

    End If

    On Error GoTo 0

    '--------------------------------
    ' insert data
    '--------------------------------
    Dim r As Long
    Dim c As Long

    ' header
    For c = 1 To m_ColCount

        ws.Cells(1, c).Value = m_Headers(c)

    Next c

    ' data
    For r = 1 To UBound(m_FiltData, 1)

        For c = 1 To m_ColCount

            ws.Cells(r + 1, c).Value = m_FiltData(r, c)

        Next c

    Next r

    '--------------------------------
    ' formatting
    '--------------------------------
    With ws.Range( _
        ws.Cells(1, 1), _
        ws.Cells(UBound(m_FiltData, 1) + 1, m_ColCount))

        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    ' header style
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, m_ColCount))

        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)

    End With

    ws.Columns.AutoFit

    MsgBox "Data successfully inserted into new sheet.", vbInformation

End Sub
'========================================================
' INSERT INTO EXISTING SHEET
'========================================================
Private Sub InsertIntoExistingSheet()

    Dim rng As Range

    On Error Resume Next

    Set rng = Application.InputBox( _
    "Select starting cell", _
    "Insert Data", _
    Type:=8)

    On Error GoTo 0

    If rng Is Nothing Then Exit Sub

    Dim r As Long
    Dim c As Long

    ' header
    For c = 1 To m_ColCount

        rng.Cells(1, c).Value = m_Headers(c)

    Next c

    ' data
    For r = 1 To UBound(m_FiltData, 1)

        For c = 1 To m_ColCount

            rng.Offset(r, c - 1).Value = m_FiltData(r, c)

        Next c

    Next r

    MsgBox "Data inserted into selected range."

End Sub


'========================================================
' DARK MODE
'========================================================

Private Sub btnDarkMode_Click()

    Dim ctrl As control

    isDarkMode = Not isDarkMode

    If isDarkMode Then

        ' DARK MODE
        Me.BackColor = RGB(35, 35, 35)

        For Each ctrl In Me.Controls

            Select Case TypeName(ctrl)

                Case "Frame"

                    ctrl.BackColor = RGB(45, 45, 45)
                    ctrl.ForeColor = vbWhite

                Case "Label"

                    ctrl.BackColor = RGB(35, 35, 35)
                    ctrl.ForeColor = vbWhite

                Case "TextBox", "ComboBox"

                    ctrl.BackColor = RGB(55, 55, 55)
                    ctrl.ForeColor = vbWhite

                Case "ListBox"

                    ctrl.BackColor = RGB(45, 45, 45)
                    ctrl.ForeColor = vbWhite

                Case "CommandButton"

                    ctrl.ForeColor = vbWhite

            End Select

        Next ctrl

        btnDarkMode.Caption = "Light Mode"

    Else

        ' LIGHT MODE
        Me.BackColor = RGB(240, 240, 240)

        For Each ctrl In Me.Controls

            Select Case TypeName(ctrl)

                Case "Frame"

                    ctrl.BackColor = RGB(240, 240, 240)
                    ctrl.ForeColor = vbBlack

                Case "Label"

                    ctrl.BackColor = RGB(240, 240, 240)
                    ctrl.ForeColor = vbBlack

                Case "TextBox", "ComboBox"

                    ctrl.BackColor = vbWhite
                    ctrl.ForeColor = vbBlack

                Case "ListBox"

                    ctrl.BackColor = vbWhite
                    ctrl.ForeColor = vbBlack

                Case "CommandButton"

                    ctrl.ForeColor = vbBlack

            End Select

        Next ctrl

        btnDarkMode.Caption = "Dark Mode"

    End If

End Sub

'========================================================
' TOP / BOTTOM ANALYSIS
'========================================================
Private Sub btnApplyTop_Click()

    If cboTopValCol.Value = "" Then

        MsgBox "Select value column."
        Exit Sub

    End If

    Dim colIdx As Long
    Dim i As Long

    For i = 1 To m_ColCount

        If m_Headers(i) = cboTopValCol.Value Then

            colIdx = i
            Exit For

        End If

    Next i

    If colIdx = 0 Then

        MsgBox "Column not found."
        Exit Sub

    End If

    Dim topCount As Long

    topCount = val(txtTopCount.Value)

    If topCount <= 0 Then topCount = 5

    Dim tempArr() As Variant

    tempArr = m_FiltData

    Dim r As Long
    Dim c As Long
    Dim tempVal As Variant

    Dim maxRows As Long

    maxRows = UBound(tempArr, 1)

    ' simple bubble sort
    For r = 1 To maxRows - 1

        For c = r + 1 To maxRows

            If cboTopN.Value = "Top" Then

                If val(tempArr(r, colIdx)) < _
                val(tempArr(c, colIdx)) Then

                    SwapRows tempArr, r, c

                End If

            Else

                If val(tempArr(r, colIdx)) > _
                val(tempArr(c, colIdx)) Then

                    SwapRows tempArr, r, c

                End If

            End If

        Next c

    Next r

    If topCount > maxRows Then
        topCount = maxRows
    End If

    ReDim m_FiltData(1 To topCount, 1 To m_ColCount)

    For r = 1 To topCount

        For c = 1 To m_ColCount

            m_FiltData(r, c) = tempArr(r, c)

        Next c

    Next r

    RenderDataList

    MsgBox cboTopN.Value & _
    " records loaded."

End Sub
'========================================================
' SWAP ROWS
'========================================================
Private Sub SwapRows(arr As Variant, row1 As Long, row2 As Long)

    Dim i As Long
    Dim tmp As Variant

    For i = 1 To m_ColCount

        tmp = arr(row1, i)

        arr(row1, i) = arr(row2, i)

        arr(row2, i) = tmp

    Next i

End Sub
'========================================================
' CONDITION FILTER
'========================================================
Private Sub btnApplyCondition_Click()

    If cboGTLTCol.Value = "" Then

        MsgBox "Select a column."
        Exit Sub

    End If

    If cboGTLTOp.Value = "" Then

        MsgBox "Select an operator."
        Exit Sub

    End If

    If Trim(txtGTLTVal.Value) = "" Then

        MsgBox "Enter condition value."
        Exit Sub

    End If

    Dim colIdx As Long
    Dim i As Long

    For i = 1 To m_ColCount

        If m_Headers(i) = cboGTLTCol.Value Then

            colIdx = i
            Exit For

        End If

    Next i

    If colIdx = 0 Then

        MsgBox "Column not found."
        Exit Sub

    End If

    Dim compareVal As Double
    Dim op As String

    compareVal = val(txtGTLTVal.Value)

    op = cboGTLTOp.Value

    Dim tempArr() As Variant

    ReDim tempArr(1 To UBound(m_OrigData, 1), 1 To m_ColCount)

    Dim r As Long
    Dim c As Long
    Dim cnt As Long

    Dim rowVal As Double
    Dim ok As Boolean

    cnt = 0

    For r = 1 To UBound(m_OrigData, 1)

        rowVal = val(m_OrigData(r, colIdx))

        ok = False

        Select Case op

            Case ">"
                ok = rowVal > compareVal

            Case "<"
                ok = rowVal < compareVal

            Case "="
                ok = rowVal = compareVal

            Case ">="
                ok = rowVal >= compareVal

            Case "<="
                ok = rowVal <= compareVal

        End Select

        If ok Then

            cnt = cnt + 1

            For c = 1 To m_ColCount

                tempArr(cnt, c) = _
                m_OrigData(r, c)

            Next c

        End If

    Next r

    If cnt = 0 Then

        lstData.Clear

        lblRowCount.Caption = _
        "Total Records: 0"

        MsgBox "No matching records found."

        Exit Sub

    End If

    ReDim m_FiltData(1 To cnt, 1 To m_ColCount)

    For r = 1 To cnt

        For c = 1 To m_ColCount

            m_FiltData(r, c) = _
            tempArr(r, c)

        Next c

    Next r

    RenderDataList

    MsgBox cnt & " records matched."

End Sub
Private Sub lstData_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = vbKeyDown Then

        If lstData.ListIndex < lstData.ListCount - 1 Then
            lstData.ListIndex = lstData.ListIndex + 1
        End If

    End If

    If KeyCode = vbKeyUp Then

        If lstData.ListIndex > 0 Then
            lstData.ListIndex = lstData.ListIndex - 1
        End If

    End If

End Sub


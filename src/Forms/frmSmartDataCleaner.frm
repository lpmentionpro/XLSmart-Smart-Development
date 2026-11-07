VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSmartDataCleaner 
   Caption         =   "SMART DATA CLEANER"
   ClientHeight    =   11820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14160
   OleObjectBlob   =   "frmSmartDataCleaner.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSmartDataCleaner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private selectedFile As String

'==================================
' USERFORM INITIALIZE
'==================================
Private Sub UserForm_Initialize()

    '-------------------------
    ' PLACEHOLDER TEXT
    '-------------------------
    txtFilePath.Value = _
        "Browse Excel file..."

    txtOutputFolder.Value = _
        "Browse export location..."

    txtFind.Value = _
        "Find value text"

    txtReplace.Value = _
        "Replace value text"

    txtSymbols.Value = _
        "Example: @#$%"


    '-------------------------
    ' TEXT CASE LIST
    '-------------------------
    cmbTextCase.Clear

    cmbTextCase.AddItem "UPPER CASE"
    cmbTextCase.AddItem "lower case"
    cmbTextCase.AddItem "Proper Case"


    '-------------------------
    ' DATE COLUMN LIST
    '-------------------------
    cmbDateColumn.Clear

    cmbDateColumn.AddItem "Joining Date"
    cmbDateColumn.AddItem "DOB"
    cmbDateColumn.AddItem "Date"
    cmbDateColumn.AddItem "Created Date"


    '-------------------------
    ' DATE FORMAT LIST
    '-------------------------
    cmbDateFormat.Clear

    cmbDateFormat.AddItem "dd-mm-yyyy"
    cmbDateFormat.AddItem "dd/mm/yyyy"
    cmbDateFormat.AddItem "dd-mm-yy"
    cmbDateFormat.AddItem "mm-dd-yyyy"
    cmbDateFormat.AddItem "mm/dd/yyyy"
    cmbDateFormat.AddItem "yyyy-mm-dd"
    cmbDateFormat.AddItem "yyyy/mm/dd"
    cmbDateFormat.AddItem "dd-mmm-yyyy"


    '-------------------------
    ' DEFAULT TEXT
    '-------------------------
    cmbTextCase.text = _
        "Select Text Case"

    cmbDateColumn.text = _
        "Select Date Column"

    cmbDateFormat.text = _
        "Select Date Format"

    btnCleanData.SetFocus

End Sub



'==================================
' BROWSE FILE
'==================================
Private Sub btnBrowse_Click()

    Dim fd As FileDialog
    Dim wb As Workbook
    Dim ws As Worksheet

    Dim lastRow As Long
    Dim lastCol As Long

    Dim r As Long
    Dim c As Long
    Dim previewCols As Long

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd

        .Title = "Select Excel File"
        .Filters.Clear
        .Filters.Add "Excel Files", _
        "*.xlsx;*.xls;*.csv"

        If .Show <> -1 Then Exit Sub

        selectedFile = .SelectedItems(1)

    End With

    txtFilePath.Value = selectedFile

    lstPreview.Clear
    cmbDateColumn.Clear

    Application.ScreenUpdating = False

    Set wb = Workbooks.Open(selectedFile, ReadOnly:=True)
    Set ws = wb.Sheets(1)

    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    ' Date columns
    For c = 1 To lastCol

        If Trim(ws.Cells(1, c).Value) <> "" Then

            cmbDateColumn.AddItem _
                ws.Cells(1, c).Value

        End If

    Next c

    ' Preview setup
    If lastCol > 10 Then
        previewCols = 10
    Else
        previewCols = lastCol
    End If

    lstPreview.Clear
    lstPreview.ColumnCount = previewCols

    Select Case previewCols

        Case 1
            lstPreview.ColumnWidths = "120 pt"

        Case 2
            lstPreview.ColumnWidths = _
                "100 pt;100 pt"

        Case 3
            lstPreview.ColumnWidths = _
                "90 pt;90 pt;90 pt"

        Case 4
            lstPreview.ColumnWidths = _
                "80 pt;80 pt;80 pt;80 pt"

        Case 5
            lstPreview.ColumnWidths = _
                "80 pt;90 pt;90 pt;90 pt;90 pt"

        Case Else

            lstPreview.ColumnWidths = _
                "90 pt;90 pt;90 pt;90 pt;90 pt;" & _
                "90 pt;90 pt;90 pt;90 pt;90 pt"

    End Select

    ' Load preview
    For r = 1 To lastRow

        lstPreview.AddItem

        For c = 1 To previewCols

            On Error Resume Next

            lstPreview.List(r - 1, c - 1) = _
                ws.Cells(r, c).text

            On Error GoTo 0

        Next c

    Next r

    wb.Close False

    Application.ScreenUpdating = True

End Sub


'==================================
' OUTPUT FOLDER
'==================================
Private Sub btnBrowseOutput_Click()

    With Application.FileDialog( _
        msoFileDialogFolderPicker)

        .Title = "Select Output Folder"

        If .Show = -1 Then

            txtOutputFolder.Value = _
                .SelectedItems(1)

        End If

    End With

End Sub


'==================================
' ADD SYMBOL
'==================================
Private Sub btnAddSymbol_Click()

    If Trim(txtSymbols.Value) <> "" Then

        lstSymbols.AddItem _
            txtSymbols.Value

        txtSymbols.Value = ""

    End If

End Sub


'==================================
' REMOVE SYMBOL
'==================================
Private Sub btnRemoveSymbol_Click()

    If lstSymbols.ListIndex >= 0 Then

        lstSymbols.RemoveItem _
            lstSymbols.ListIndex

    End If

End Sub


'==================================
' SELECT ALL
'==================================
Private Sub chkSelectAll_Click()

    Dim ctrl As control

    For Each ctrl In Me.Controls

        If TypeName(ctrl) = "CheckBox" Then

            ctrl.Value = _
                chkSelectAll.Value

        End If

    Next ctrl

End Sub



Private Sub btnCleanData_Click()

   '=========================
    ' CHECK OUTPUT FOLDER
    '=========================
    If Trim(txtOutputFolder.Value) = "" Then

    MsgBox "Please select output folder first.", _
           vbExclamation, _
           "Output Folder Missing"

    txtOutputFolder.SetFocus
    Exit Sub

End If

    Dim wbSource As Workbook
    Dim wsSource As Worksheet

    Dim wbClean As Workbook
    Dim wsClean As Worksheet

    Dim cell As Range
    Dim rng As Range

    Dim rw As Long
    Dim cc As Long
    Dim i As Long
    Dim itemNo As Long

    Dim lastCol As Long
    Dim colIndex As Long

    Dim sym As String
    Dim savePath As String
    Dim dupCols() As Variant

    If Trim(txtFilePath.Value) = "" Then

        MsgBox "Please select a source file first.", vbExclamation
        Exit Sub

    End If

    ' create output folder
    If Dir(txtOutputFolder.Value, vbDirectory) = "" Then
        MkDir txtOutputFolder.Value
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wbSource = Workbooks.Open(txtFilePath.Value)
    Set wsSource = wbSource.Sheets(1)

    Set wbClean = Workbooks.Add
    Set wsClean = wbClean.Sheets(1)

    wsClean.Name = "Cleaned Data"

    wsSource.UsedRange.Copy
    wsClean.Range("A1").PasteSpecial xlPasteAll

    Application.CutCopyMode = False

    lastCol = wsClean.Cells(1, wsClean.Columns.count).End(xlToLeft).Column

    '========================
    ' TRIM EXTRA SPACES
    '========================
    If chkTrimSpaces.Value = True Then

        For Each cell In wsClean.UsedRange

            If VarType(cell.Value) = vbString Then
                cell.Value = Application.Trim(cell.Value)
            End If

        Next cell

    End If


    '========================
    ' TEXT CASE
    '========================
    If chkTextCase.Value = True Then

        For Each cell In wsClean.UsedRange

            If VarType(cell.Value) = vbString Then

                Select Case cmbTextCase.Value

                    Case "UPPER CASE"
                        cell.Value = UCase(cell.Value)

                    Case "lower case"
                        cell.Value = LCase(cell.Value)

                    Case "Proper Case"
                        cell.Value = StrConv(cell.Value, vbProperCase)

                End Select

            End If

        Next cell

    End If


    '========================
    ' REMOVE HYPERLINKS
    '========================
    If chkRemoveLinks.Value = True Then

        On Error Resume Next
        wsClean.Hyperlinks.Delete
        On Error GoTo 0

    End If


    '========================
    ' FORMULA TO VALUES
    '========================
    If chkFormulaToValues.Value = True Then

        wsClean.UsedRange.Value = _
            wsClean.UsedRange.Value

    End If



    '========================
' REMOVE SYMBOLS
'========================
If chkSymbols.Value = True Then

    Dim fullText As String
    Dim singleChar As String
    Dim x As Long

    If lstSymbols.ListCount > 0 Then

        For itemNo = 0 To _
            lstSymbols.ListCount - 1

            fullText = Trim( _
                lstSymbols.List(itemNo) & "")

            If fullText <> "" Then

                ' @#$% ??? ?? ?? ??? remove ????
                For x = 1 To Len(fullText)

                    singleChar = _
                        Mid(fullText, x, 1)

                    wsClean.Cells.Replace _
                        What:=singleChar, _
                        replacement:="", _
                        LookAt:=xlPart, _
                        SearchOrder:=xlByRows, _
                        MatchCase:=False

                Next x

            End If

        Next itemNo

    End If

End If


'=========================
' FIND & REPLACE
'=========================
If chkFindReplace.Value Then

    Dim searchText As String
    Dim replaceText As String
    Dim txtCell As String

    searchText = Trim(txtFind.text)
    replaceText = Trim(txtReplace.text)

    ' nothing typed
    If searchText <> "" Then

        '=====================
        ' WHOLE WORD ONLY
        '=====================
        If chkWholeWordOnly.Value = True Then

            For Each cell In wsClean.UsedRange

                If Not IsError(cell.Value) Then

                    txtCell = Trim(CStr(cell.Value))

                    ' exact match only
                    If StrComp(txtCell, _
                               searchText, _
                               vbTextCompare) = 0 Then

                        cell.Value = replaceText

                    End If

                End If

            Next cell

        Else

            '=====================
            ' PARTIAL MATCH
            '=====================
            wsClean.Cells.Replace _
                What:=searchText, _
                replacement:=replaceText, _
                LookAt:=xlPart, _
                SearchOrder:=xlByRows, _
                MatchCase:=False

        End If

    End If

End If

    '========================
    ' REMOVE DUPLICATES
    '========================
    If chkDuplicates.Value = True Then

        ReDim dupCols(1 To lastCol)

        For i = 1 To lastCol
            dupCols(i) = i
        Next i

        On Error Resume Next

        wsClean.UsedRange.RemoveDuplicates _
            Columns:=dupCols, _
            Header:=xlYes

        On Error GoTo 0

    End If


    '========================
    ' DELETE EMPTY ROWS
    '========================
    If chkEmptyRows.Value = True Then

        For rw = wsClean.Cells( _
            wsClean.Rows.count, 1) _
            .End(xlUp).Row To 2 Step -1

            If WorksheetFunction.CountA( _
                wsClean.Rows(rw)) = 0 Then

                wsClean.Rows(rw).Delete

            End If

        Next rw

    End If


    '========================
    ' DELETE EMPTY COLUMNS
    '========================
    If chkEmptyColumns.Value = True Then

        For cc = wsClean.Cells( _
            1, wsClean.Columns.count) _
            .End(xlToLeft).Column To 1 Step -1

            If WorksheetFunction.CountA( _
                wsClean.Columns(cc)) = 0 Then

                wsClean.Columns(cc).Delete

            End If

        Next cc

    End If


    '========================
    ' FILL TEXT BLANKS
    '========================
    If chkFillTextNA.Value = True Then

        For Each rng In wsClean.UsedRange

            If Trim(rng.text) = "" Then
                rng.Value = "N/A"
            End If

        Next rng

    End If


    '========================
    ' FILL NUMERIC BLANKS
    '========================
    If chkFillZero.Value = True Then

        For Each rng In wsClean.UsedRange

            If Trim(rng.text) = "" Then
                rng.Value = 0
            End If

        Next rng

    End If


    '========================
    ' CLEAR COLORS & FORMATTING
    '========================
    If chkClearFormatting.Value = True Then

        wsClean.Cells.ClearFormats

    End If


    '========================
    ' REMOVE COMMENTS
    '========================
    If chkRemoveComments.Value = True Then

        On Error Resume Next
        wsClean.Cells.ClearComments
        wsClean.Cells.ClearNotes
        On Error GoTo 0

    End If


    '========================
    ' DELETE OBJECTS / SHAPES
    '========================
    If chkDeleteShapes.Value = True Then

        On Error Resume Next

        Dim shp As Shape

        For Each shp In wsClean.Shapes
            shp.Delete
        Next shp

        On Error GoTo 0

    End If


    '========================
    ' FIX DATE FORMAT
    '========================
    If chkDateFormat.Value = True Then

        colIndex = 0

        For i = 1 To lastCol

            If Trim(wsClean.Cells(1, i).Value) = _
                Trim(cmbDateColumn.Value) Then

                colIndex = i
                Exit For

            End If

        Next i

        If colIndex > 0 Then

            wsClean.Columns(colIndex) _
                .NumberFormat = _
                cmbDateFormat.Value

        End If

    End If


    wsClean.Columns.AutoFit


'-----------------------
' SAVE FILE
'-----------------------
Dim outputFolder As String

outputFolder = Trim(txtOutputFolder.Value)

If outputFolder = "" Then
    outputFolder = "C:\CleanedData"
End If

' folder ?? ????? create
If Dir(outputFolder, vbDirectory) = "" Then
    MkDir outputFolder
End If

savePath = outputFolder & _
            "\Cleaned_" & _
            Format(Now, "yyyymmdd_hhnnss") & _
            ".xlsx"

wbClean.SaveAs Filename:=savePath, _
               FileFormat:=xlOpenXMLWorkbook

    wbSource.Close False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Data cleaned successfully!", _
        vbInformation

End Sub


'==================================
' RESET BUTTON
'==================================
'==================================
' RESET BUTTON
'==================================
Private Sub btnReset_Click()

    Dim ctrl As control

    '-------------------------
    ' UNCHECK ALL CHECKBOX
    '-------------------------
    For Each ctrl In Me.Controls

        If TypeName(ctrl) = "CheckBox" Then

            ctrl.Value = False

        End If

    Next ctrl


    '-------------------------
    ' RESET TEXTBOX
    '-------------------------
    txtFilePath.Value = _
        "Browse Excel file..."

    txtOutputFolder.Value = _
        "Browse export location..."

    txtFind.Value = _
        "Find value text"

    txtReplace.Value = _
        "Replace value text"

    txtSymbols.Value = _
        "Example: @#$%"


    '-------------------------
    ' RESET COMBOBOX
    '-------------------------
    cmbTextCase.ListIndex = -1
    cmbDateColumn.ListIndex = -1
    cmbDateFormat.ListIndex = -1

    cmbTextCase.text = _
        "Select Text Case"

    cmbDateColumn.text = _
        "Select Date Column"
    
    
    '-------------------------
' DATE FORMAT LIST
'-------------------------
cmbDateFormat.Clear

cmbDateFormat.AddItem "dd-mm-yyyy"
cmbDateFormat.AddItem "dd/mm/yyyy"
cmbDateFormat.AddItem "dd-mm-yy"
cmbDateFormat.AddItem "mm-dd-yyyy"
cmbDateFormat.AddItem "mm/dd/yyyy"
cmbDateFormat.AddItem "yyyy-mm-dd"
cmbDateFormat.AddItem "yyyy/mm/dd"
cmbDateFormat.AddItem "dd-mmm-yyyy"
cmbDateFormat.AddItem "mmm-dd-yyyy"

    cmbDateFormat.text = _
        "Select Date Format"


    '-------------------------
    ' CLEAR LIST
    '-------------------------
    lstSymbols.Clear
    lstPreview.Clear


    '-------------------------
    ' SELECT ALL OFF
    '-------------------------
    chkSelectAll.Value = False

End Sub


'==================================
' AUTO CLEAR PLACEHOLDER
'==================================

Private Sub txtFind_Enter()

    If txtFind.Value = _
        "Find value text" Then

        txtFind.Value = ""

    End If

End Sub


Private Sub txtReplace_Enter()

    If txtReplace.Value = _
        "Replace value text" Then

        txtReplace.Value = ""

    End If

End Sub


Private Sub txtSymbols_Enter()

    If txtSymbols.Value = _
        "Example: @#$%" Then

        txtSymbols.Value = ""

    End If

End Sub


Private Sub txtFilePath_Enter()

    If txtFilePath.Value = _
        "Browse Excel file..." Then

        txtFilePath.Value = ""

    End If

End Sub

Private Sub txtOutputFolder_Enter()

    If txtOutputFolder.Value = _
        "Browse export location..." Then

        txtOutputFolder.Value = ""

    End If

End Sub


'==================================
' OPEN FOLDER
'==================================
Private Sub btnOpenOutputFolder_Click()

    If Dir(txtOutputFolder.Value, _
       vbDirectory) <> "" Then

        Shell "explorer.exe """ & _
        txtOutputFolder.Value & """", _
        vbNormalFocus

    Else

        MsgBox "Folder not found."

    End If

End Sub

'==================================
' DARK MODE
'==================================
Private Sub btnDarkMode_Click()

    Static isDark As Boolean
    Dim ctrl As control

    isDark = Not isDark

    If isDark Then

        ' DARK MODE
        Me.BackColor = RGB(35, 35, 35)

        For Each ctrl In Me.Controls

            Select Case TypeName(ctrl)

                Case "Label"

                    ctrl.ForeColor = RGB(255, 255, 255)
                    ctrl.BackStyle = fmBackStyleTransparent

                Case "Frame"

                    ctrl.BackColor = RGB(45, 45, 45)
                    ctrl.ForeColor = RGB(255, 255, 255)

                Case "TextBox"

                    ctrl.BackColor = RGB(55, 55, 55)
                    ctrl.ForeColor = RGB(255, 255, 255)

                Case "ComboBox"

                    ctrl.BackColor = RGB(55, 55, 55)
                    ctrl.ForeColor = RGB(255, 255, 255)

                Case "ListBox"

                    ctrl.BackColor = RGB(55, 55, 55)
                    ctrl.ForeColor = RGB(255, 255, 255)

                Case "CheckBox"

                    ctrl.ForeColor = RGB(255, 255, 255)
                    ctrl.BackColor = RGB(35, 35, 35)

                Case "CommandButton"

                    ctrl.ForeColor = RGB(255, 255, 255)

            End Select

        Next ctrl

        btnDarkMode.Caption = "Light Mode"

    Else

        ' LIGHT MODE
        Me.BackColor = RGB(242, 242, 242)

        For Each ctrl In Me.Controls

            Select Case TypeName(ctrl)

                Case "Label"

                    ctrl.ForeColor = RGB(0, 0, 0)
                    ctrl.BackStyle = fmBackStyleTransparent

                Case "Frame"

                    ctrl.BackColor = RGB(242, 242, 242)
                    ctrl.ForeColor = RGB(0, 0, 0)

                Case "TextBox"

                    ctrl.BackColor = RGB(255, 255, 255)
                    ctrl.ForeColor = RGB(0, 0, 0)

                Case "ComboBox"

                    ctrl.BackColor = RGB(255, 255, 255)
                    ctrl.ForeColor = RGB(0, 0, 0)

                Case "ListBox"

                    ctrl.BackColor = RGB(255, 255, 255)
                    ctrl.ForeColor = RGB(0, 0, 0)

                Case "CheckBox"

                    ctrl.ForeColor = RGB(0, 0, 0)
                    ctrl.BackColor = RGB(242, 242, 242)

                Case "CommandButton"

                    ctrl.ForeColor = RGB(0, 0, 0)

            End Select

        Next ctrl

        btnDarkMode.Caption = "Dark Mode"

    End If

End Sub


'==================================
' CLOSE
'==================================
Private Sub btnCancel_Click()

    Unload Me

End Sub


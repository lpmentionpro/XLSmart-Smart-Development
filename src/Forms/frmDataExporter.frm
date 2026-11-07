VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataExporter 
   Caption         =   "Smart Bulk Exporter"
   ClientHeight    =   10215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7515
   OleObjectBlob   =   "frmDataExporter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDataExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'=======================================================
' USERFORM : frmDataExporter
' Smart Bulk Exporter - Final Professional Version
'=======================================================

Option Explicit

Private m_DataRange As Range
Private m_UniqueItems As Collection

'=======================================================
' INIT FORM
'=======================================================

Public Sub InitForm(srcRange As Range)

    Set m_DataRange = srcRange

    Dim i As Long

    cboSplitCol.Clear
    cboSplitCol.AddItem "Select Column"

    For i = 1 To srcRange.Columns.count

        cboSplitCol.AddItem _
        CStr(srcRange.Cells(1, i).Value)

    Next i

    cboSplitCol.ListIndex = 0

    lstUniqueItems.Clear
    txtSearchUniq.text = "Search Unique Items..."
    txtSearchUniq.ForeColor = RGB(120, 120, 120)

End Sub

Private Sub txtSearchUniq_Enter()

    If txtSearchUniq.text = "Search Unique Items..." Then

        txtSearchUniq.text = ""
        txtSearchUniq.ForeColor = vbBlack

    End If

End Sub

Private Sub txtSearchUniq_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Trim(txtSearchUniq.text) = "" Then

        txtSearchUniq.text = "Search Unique Items..."
        txtSearchUniq.ForeColor = RGB(120, 120, 120)

    End If

End Sub

'=======================================================
' COLUMN CHANGE
'=======================================================

Private Sub cboSplitCol_Change()

    If cboSplitCol.ListIndex <= 0 Then Exit Sub

    lstUniqueItems.Clear

    Set m_UniqueItems = New Collection

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim colIdx As Long
    colIdx = cboSplitCol.ListIndex

    Dim dataCol As Range
    Dim cell As Range

    Set dataCol = m_DataRange.Columns(colIdx)

    For Each cell In dataCol.Cells

        ' skip header
        If cell.Row > dataCol.Cells(1).Row Then

            Dim cleanVal As String

            cleanVal = Trim(CStr(cell.Value))

            cleanVal = WorksheetFunction.Trim(cleanVal)

            cleanVal = Replace(cleanVal, Chr(160), "")

            If cleanVal <> "" Then

                If Not dict.exists(LCase(cleanVal)) Then

                    dict.Add LCase(cleanVal), cleanVal

                    m_UniqueItems.Add cleanVal

                    lstUniqueItems.AddItem cleanVal

                End If

            End If

        End If

    Next cell

End Sub

Private Sub Frame2_Click()

End Sub

'=======================================================
' SEARCH
'=======================================================
Private Sub txtSearchUniq_Change()

    If m_UniqueItems Is Nothing Then Exit Sub

    If txtSearchUniq.text = "Search Unique Items..." Then Exit Sub

    Dim txt As String
    txt = LCase(Trim(txtSearchUniq.text))

    lstUniqueItems.Clear

    Dim itm As Variant

    For Each itm In m_UniqueItems

        If txt = "" Then

            lstUniqueItems.AddItem itm

        Else

            If InStr(1, LCase(itm), txt) > 0 Then

                lstUniqueItems.AddItem itm

            End If

        End If

    Next itm

End Sub


'=======================================================
' REMOVE ITEM
'=======================================================
Private Sub btnRemoveItem_Click()

    Dim searchVal As String

    searchVal = LCase(Trim(txtSearchUniq.text))

    If searchVal = "" Then

        MsgBox "Type something to remove.", vbExclamation
        Exit Sub

    End If

    If searchVal = "search unique items..." Then Exit Sub

    Dim i As Long

    For i = lstUniqueItems.ListCount - 1 To 0 Step -1

        If InStr(1, _
           LCase(lstUniqueItems.List(i)), _
           searchVal) > 0 Then

            lstUniqueItems.RemoveItem i

        End If

    Next i

    txtSearchUniq.text = ""

End Sub


'=======================================================
' EXPORT PDF
'=======================================================

Private Sub btnExportPDF_Click()

    If lstUniqueItems.ListCount > 50 Then

        MsgBox _
        "Unique items exceeded 50." & vbCrLf & vbCrLf & _
        "Please use Search + Remove button " & _
        "to reduce the list before export.", _
        vbExclamation

        Exit Sub

    End If

    ExportFilteredData "PDF"

End Sub

'=======================================================
' EXPORT EXCEL
'=======================================================
Private Sub btnExportExcel_Click()

    If lstUniqueItems.ListCount > 50 Then

        MsgBox _
        "Unique items exceeded 50." & vbCrLf & vbCrLf & _
        "Please use Search + Remove button " & _
        "to reduce the list before export.", _
        vbExclamation

        Exit Sub

    End If

    ExportFilteredData "EXCEL"

End Sub

'=======================================================
' EXPORT MAIN
'=======================================================

Private Sub ExportFilteredData(exportType As String)

    If cboSplitCol.ListIndex <= 0 Then

        MsgBox "Select a column.", vbExclamation
        Exit Sub

    End If

    If lstUniqueItems.ListCount = 0 Then

        MsgBox "No unique items found.", vbExclamation
        Exit Sub

    End If

    Dim fd As FileDialog
    Dim savePath As String
    Dim exportFolder As String

    Set fd = Application.FileDialog(msoFileDialogSaveAs)

    With fd

        .Title = "Select Export Location"

        .InitialFileName = _
        ThisWorkbook.Path & "\Exported_Data"

        If .Show <> -1 Then Exit Sub

        savePath = .SelectedItems(1)

    End With

    exportFolder = left(savePath, _
    InStrRev(savePath, "\") - 1)


    Dim wsSource As Worksheet
    Set wsSource = m_DataRange.Worksheet

    Dim colIdx As Long
    colIdx = cboSplitCol.ListIndex

    Dim i As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For i = 0 To lstUniqueItems.ListCount - 1

        Dim uniqueVal As String
        uniqueVal = lstUniqueItems.List(i)

        Application.StatusBar = _
        "Exporting " & (i + 1) & _
        " of " & lstUniqueItems.ListCount & _
        " : " & uniqueVal

        m_DataRange.AutoFilter _
            Field:=colIdx, _
            Criteria1:=uniqueVal

        Dim visibleRng As Range

        On Error Resume Next

        Set visibleRng = _
        m_DataRange.SpecialCells(xlCellTypeVisible)

        On Error GoTo 0

        If visibleRng Is Nothing Then GoTo skipExport

        Dim wb As Workbook
        Set wb = Workbooks.Add

        Dim ws As Worksheet
        Set ws = wb.Sheets(1)

        visibleRng.Copy

        ws.Range("A1").PasteSpecial xlPasteAll
        ws.Range("A1").PasteSpecial xlPasteColumnWidths

        Application.CutCopyMode = False

        ws.Cells.EntireColumn.AutoFit

        With ws.PageSetup

            .Orientation = xlLandscape
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False

            .LeftMargin = _
            Application.InchesToPoints(0.3)

            .RightMargin = _
            Application.InchesToPoints(0.3)

            .TopMargin = _
            Application.InchesToPoints(0.5)

            .BottomMargin = _
            Application.InchesToPoints(0.5)

            .CenterHorizontally = True

        End With

        Dim saveName As String

        saveName = exportFolder & "\" & _
        CleanFileName(uniqueVal)

        If exportType = "PDF" Then

            ws.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                Filename:=saveName & ".pdf", _
                Quality:=xlQualityStandard

            wb.Close False

        Else

            wb.SaveAs _
                Filename:=saveName & ".xlsx", _
                FileFormat:=xlOpenXMLWorkbook

            wb.Close False

        End If

skipExport:

        Set visibleRng = Nothing

    Next i

    On Error Resume Next

    wsSource.ShowAllData

    Application.StatusBar = False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox _
    exportType & " Export Completed Successfully.", _
    vbInformation

End Sub

'=======================================================
' INSERT SHEETS
'=======================================================

Private Sub btnInsertSheets_Click()
    If lstUniqueItems.ListCount > 50 Then

    MsgBox _
    "Unique items exceeded 50." & vbCrLf & vbCrLf & _
    "Please use Search + Remove button " & _
    "to reduce the list before creating sheets.", _
    vbExclamation

    Exit Sub

End If

    If cboSplitCol.ListIndex <= 0 Then Exit Sub

    Dim colIdx As Long
    colIdx = cboSplitCol.ListIndex

    Dim wsSource As Worksheet
    Set wsSource = m_DataRange.Worksheet

    Dim i As Long

    Application.ScreenUpdating = False

    For i = 0 To lstUniqueItems.ListCount - 1

        Dim uniqueVal As String
        uniqueVal = lstUniqueItems.List(i)

        Application.StatusBar = _
        "Creating Sheet : " & uniqueVal

        m_DataRange.AutoFilter _
            Field:=colIdx, _
            Criteria1:=uniqueVal

        Dim visibleRng As Range

        On Error Resume Next

        Set visibleRng = _
        m_DataRange.SpecialCells(xlCellTypeVisible)

        On Error GoTo 0

        If visibleRng Is Nothing Then GoTo skipSheet

        Dim newWs As Worksheet

        Set newWs = Sheets.Add(After:=Sheets(Sheets.count))

        On Error Resume Next

        newWs.Name = _
        left(CleanFileName(uniqueVal), 31)

        On Error GoTo 0

        visibleRng.Copy

        newWs.Range("A1").PasteSpecial xlPasteAll
        newWs.Range("A1").PasteSpecial xlPasteColumnWidths

        Application.CutCopyMode = False

        newWs.Cells.EntireColumn.AutoFit

skipSheet:

        Set visibleRng = Nothing

    Next i

    On Error Resume Next

    wsSource.ShowAllData

    Application.StatusBar = False

    Application.ScreenUpdating = True

    MsgBox _
    "Sheets Created Successfully.", vbInformation

End Sub

'=======================================================
' CLEAN FILE NAME
'=======================================================

Private Function CleanFileName(f As String) As String

    Dim badChars As String

    badChars = "\/:*?""<>|"

    Dim i As Long

    For i = 1 To Len(badChars)

        f = Replace(f, Mid(badChars, i, 1), "_")

    Next i

    CleanFileName = Trim(f)

End Function




Attribute VB_Name = "modWorkbookIndex"
' ============================================================
' GENERATE SHEET INDEX
' ============================================================
Public Sub GenerateSheetIndex(control As IRibbonControl)

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error Resume Next
    wb.Sheets("Sheet_Index").Delete
    On Error GoTo 0

    Application.DisplayAlerts = True

    Dim idx As Worksheet
    Set idx = wb.Sheets.Add(Before:=wb.Sheets(1))

    idx.Name = "Sheet_Index"

    Dim totalSheets As Integer
    totalSheets = wb.Sheets.count - 1

    ' top spacing
    idx.Rows(1).RowHeight = 34
    idx.Rows(2).RowHeight = 20
    idx.Rows(3).RowHeight = 24

    ' =========================================================
    ' HEADER
    ' =========================================================

    With idx.Range("A1:G1")

        .Merge

        .Value = "SHEET INDEX DASHBOARD  •  " & totalSheets & " SHEETS"

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .Font.Name = "Segoe UI"
        .Font.Size = 15
        .Font.Bold = True

        .Font.Color = RGB(255, 255, 255)

        .Interior.Color = RGB(36, 74, 140)

    End With

    With idx.Range("A2:G2")

        .Merge

        .Value = "Click sheet name to open  •  Use delete button to remove sheet"

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Italic = True

        .Font.Color = RGB(90, 90, 90)

        .Interior.Color = RGB(240, 244, 250)

    End With

    ' =========================================================
    ' TABLE HEADER
    ' =========================================================

    Dim hdr As Variant

    hdr = Array( _
        "No", _
        "Sheet Name", _
        "Used Rows", _
        "Used Cols", _
        "Protection", _
        "Visibility", _
        "Action")

    Dim i As Integer

    For i = 0 To 6

        With idx.Cells(3, i + 1)

            .Value = hdr(i)

            .Font.Name = "Segoe UI"
            .Font.Size = 10
            .Font.Bold = True

            .Font.Color = RGB(255, 255, 255)

            .Interior.Color = RGB(56, 96, 170)

            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter

        End With

    Next i

    ' =========================================================
    ' DATA ROWS
    ' =========================================================

    Dim dataRow As Long
    dataRow = 4

    Dim serialNo As Integer
    serialNo = 1

    Dim ws As Worksheet

    For Each ws In wb.Worksheets

        If ws.Name = "Sheet_Index" Then GoTo skipSheet

        idx.Rows(dataRow).RowHeight = 22

        Dim bg As Long

        If serialNo Mod 2 = 0 Then
            bg = RGB(245, 247, 252)
        Else
            bg = RGB(255, 255, 255)
        End If

        ' ----------------------------------------
        ' NO
        ' ----------------------------------------

        With idx.Cells(dataRow, 1)

            .Value = serialNo

            .Interior.Color = bg

            .Font.Name = "Segoe UI"
            .Font.Size = 10

            .HorizontalAlignment = xlCenter

        End With

        ' ----------------------------------------
        ' SHEET NAME
        ' ----------------------------------------

        idx.Cells(dataRow, 2).Interior.Color = bg

        idx.Hyperlinks.Add _
            Anchor:=idx.Cells(dataRow, 2), _
            Address:="", _
            SubAddress:="'" & ws.Name & "'!A1", _
            TextToDisplay:=ws.Name

        With idx.Cells(dataRow, 2)

            .Font.Name = "Segoe UI"
            .Font.Size = 10

            .Font.Color = RGB(0, 70, 140)

        End With

        ' ----------------------------------------
        ' USED RANGE
        ' ----------------------------------------

        Dim totalRows As Long
        Dim totalCols As Long

        On Error Resume Next

        totalRows = ws.UsedRange.Rows.count
        totalCols = ws.UsedRange.Columns.count

        On Error GoTo 0

        With idx.Cells(dataRow, 3)

            .Value = totalRows

            .Interior.Color = bg

            .HorizontalAlignment = xlCenter

        End With

        With idx.Cells(dataRow, 4)

            .Value = totalCols

            .Interior.Color = bg

            .HorizontalAlignment = xlCenter

        End With

        ' ----------------------------------------
        ' PROTECTION
        ' ----------------------------------------

        With idx.Cells(dataRow, 5)

       .Interior.Color = bg
       .HorizontalAlignment = xlCenter

         If ws.ProtectContents Then

        .Value = "Protected"
        .Font.Color = RGB(190, 0, 0)

    Else

        .Value = "Unprotected"
        .Font.Color = RGB(0, 128, 0)

    End If

End With

        ' ----------------------------------------
        ' VISIBILITY
        ' ----------------------------------------

        With idx.Cells(dataRow, 6)

            .Interior.Color = bg

            .HorizontalAlignment = xlCenter

            If ws.Visible = xlSheetVisible Then

                .Value = "Visible"
                .Font.Color = RGB(0, 128, 0)

            Else

                .Value = "Hidden"
                .Font.Color = RGB(190, 0, 0)

            End If

        End With

        ' ----------------------------------------
        ' ACTION BG
        ' ----------------------------------------

        idx.Cells(dataRow, 7).Interior.Color = bg

        ' hidden sheet name
        idx.Cells(dataRow, 7).Value = ws.Name
        idx.Cells(dataRow, 7).Font.Color = bg

        ' ----------------------------------------
        ' DELETE BUTTON
        ' ----------------------------------------

        Dim delBtn As Object

        Set delBtn = idx.Buttons.Add( _
            idx.Cells(dataRow, 7).left + 2, _
            idx.Cells(dataRow, 7).Top + 1, _
            idx.Cells(dataRow, 7).Width - 4, _
            idx.Cells(dataRow, 7).Height - 2)

        With delBtn

            .Caption = "Delete"

            .Name = "DelBtn_" & serialNo

            .Font.Size = 9
            .Font.Bold = True

            .OnAction = "DeleteSheetFromIndex"

        End With

        dataRow = dataRow + 1
        serialNo = serialNo + 1

skipSheet:
    Next ws

    ' =========================================================
    ' COLUMN WIDTH
    ' =========================================================

    idx.Columns(1).ColumnWidth = 6
    idx.Columns(2).ColumnWidth = 30
    idx.Columns(3).ColumnWidth = 13
    idx.Columns(4).ColumnWidth = 13
    idx.Columns(5).ColumnWidth = 16
    idx.Columns(6).ColumnWidth = 14
    idx.Columns(7).ColumnWidth = 12

    ' =========================================================
    ' BORDER
    ' =========================================================

    With idx.Range( _
        idx.Cells(3, 1), _
        idx.Cells(dataRow - 1, 7))

        .Borders.LineStyle = xlContinuous

        .Borders.Color = RGB(210, 220, 235)

        .Borders.Weight = xlThin

    End With

    idx.Activate

    idx.Range("A4").Select

    ActiveWindow.DisplayGridlines = False

    Application.ScreenUpdating = True

    MsgBox "Sheet Index created successfully!", vbInformation

End Sub
' ============================================================
' GENERATE SHEET GALLERY
' ============================================================
Public Sub GenerateSheetGallery(control As IRibbonControl)

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error Resume Next
    wb.Sheets("Sheet_Gallery").Delete
    On Error GoTo 0

    Application.DisplayAlerts = True

    Dim gal As Worksheet
    Set gal = wb.Sheets.Add(Before:=wb.Sheets(1))

    gal.Name = "Sheet_Gallery"

    Dim c As Integer

    gal.Columns(1).ColumnWidth = 2

    For c = 2 To 100
        gal.Columns(c).ColumnWidth = 5.5
    Next c

    gal.Rows(1).RowHeight = 10
    gal.Rows(2).RowHeight = 40
    gal.Rows(3).RowHeight = 18
    gal.Rows(4).RowHeight = 8

    ' =====================================================
    ' HEADER
    ' =====================================================

    With gal.Range("B2:AO2")

        .Merge
        .Value = "WORKBOOK SHEET GALLERY"

        .Font.Name = "Segoe UI"
        .Font.Size = 18
        .Font.Bold = True

        .Font.Color = RGB(35, 35, 35)

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .Interior.Color = RGB(236, 241, 255)

    End With

    With gal.Range("B3:AO3")

        .Merge
        .Value = "Click Open button to navigate any sheet"

        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Italic = True

        .Font.Color = RGB(120, 120, 145)

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .Interior.Color = RGB(236, 241, 255)

    End With

    With gal.Range("B3:AO3").Borders(xlEdgeBottom)

        .LineStyle = xlContinuous
        .Color = RGB(190, 200, 225)
        .Weight = xlThin

    End With

    ' =====================================================
    ' CARD SETTINGS
    ' =====================================================

    Dim cardStartCols As Variant
    cardStartCols = Array(2, 16, 30)

    Const CARD_COLS As Integer = 12
    Const PREV_ROWS As Integer = 8
    Const CARD_ROWS As Integer = 12
    Const ROW_GAP As Integer = 2
    Const DATA_START As Integer = 5

    Dim cardCount As Integer
    cardCount = 0

    Dim ws As Worksheet

    For Each ws In wb.Sheets

        If ws.Name = "Sheet_Gallery" _
        Or ws.Name = "Sheet_Index" Then GoTo skipSheet

        Dim col0 As Integer
        Dim row0 As Integer
        Dim col1 As Integer

        col0 = cardStartCols(cardCount Mod 3)

        row0 = DATA_START + _
               (cardCount \ 3) * (CARD_ROWS + ROW_GAP)

        col1 = col0 + CARD_COLS - 1

        ' card bg
        With gal.Range( _
            gal.Cells(row0, col0), _
            gal.Cells(row0 + CARD_ROWS - 1, col1))

            .Interior.Color = RGB(252, 253, 255)

        End With

        ' outer border
        With gal.Range( _
            gal.Cells(row0, col0), _
            gal.Cells(row0 + CARD_ROWS - 1, col1))

            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous

            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeRight).Weight = xlMedium
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeBottom).Weight = xlMedium

            .Borders.Color = RGB(185, 195, 225)

        End With

        ' top blue bar
        With gal.Range( _
            gal.Cells(row0, col0), _
            gal.Cells(row0 + 1, col1))

            .Merge
            .Interior.Color = RGB(92, 130, 255)

        End With

        With gal.Cells(row0, col0)

            .Value = ws.Name

            .Font.Name = "Segoe UI"
            .Font.Size = 8
            .Font.Bold = True

            .Font.Color = RGB(255, 255, 255)

            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter

            .IndentLevel = 1

        End With

        ' preview
        Dim pr As Integer
        Dim pc As Integer

        For pr = 1 To PREV_ROWS

            For pc = 1 To CARD_COLS

                Dim cellVal As String
                cellVal = ""

                On Error Resume Next

                If pr <= ws.UsedRange.Rows.count _
                And pc <= ws.UsedRange.Columns.count Then

                    cellVal = left(CStr(ws.Cells(pr, pc).Value), 6)

                End If

                On Error GoTo 0

                With gal.Cells( _
                    row0 + 2 + pr - 1, _
                    col0 + pc - 1)

                    .Value = cellVal

                    .Font.Name = "Segoe UI"
                    .Font.Size = 6

                    .Font.Color = RGB(70, 70, 95)

                    If pr Mod 2 = 0 Then
                        .Interior.Color = RGB(247, 249, 255)
                    Else
                        .Interior.Color = RGB(255, 255, 255)
                    End If

                End With

            Next pc

        Next pr

        ' preview border
        With gal.Range( _
            gal.Cells(row0 + 2, col0), _
            gal.Cells(row0 + 2 + PREV_ROWS - 1, col1))

            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(222, 228, 240)
            .Borders.Weight = xlThin

        End With

        ' =================================================
        ' OPEN BUTTON
        ' =================================================

        gal.Rows(row0 + CARD_ROWS - 1).RowHeight = 24

        Dim shpBtn As Shape

        Dim btnLeft As Double
        Dim btnTop As Double
        Dim btnWidth As Double
        Dim btnHt As Double

        btnLeft = gal.Cells( _
                    row0 + CARD_ROWS - 1, col0).left + 6

        btnTop = gal.Cells( _
                   row0 + CARD_ROWS - 1, col0).Top - 2

        btnWidth = gal.Cells(row0, col1).left + _
                   gal.Columns(col1).Width - _
                   gal.Cells(row0, col0).left - 12

        btnHt = 28

        Set shpBtn = gal.Shapes.AddShape( _
                        msoShapeRoundedRectangle, _
                        btnLeft, btnTop, _
                        btnWidth, btnHt)

        With shpBtn

            .Name = "btn_" & ws.Index

            .AlternativeText = ws.Name

            .Fill.ForeColor.RGB = RGB(70, 120, 240)

            .Line.Visible = msoFalse

            .Adjustments.Item(1) = 0.25

            .Shadow.Visible = msoTrue

            .TextFrame2.TextRange.text = "Open"

            .TextFrame2.TextRange.Font.Size = 10
            .TextFrame2.TextRange.Font.Bold = msoTrue

            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                RGB(255, 255, 255)

            .TextFrame2.VerticalAnchor = msoAnchorMiddle

            .TextFrame2.TextRange.ParagraphFormat.Alignment = _
                msoAlignCenter

            .OnAction = "'" & ThisWorkbook.Name & "'!OpenSheetFromBtn"

        End With

        cardCount = cardCount + 1

skipSheet:
    Next ws

    gal.Activate

    ActiveWindow.Zoom = 100

    Application.ScreenUpdating = True

    MsgBox "Gallery ready!", vbInformation

End Sub


' ============================================================
' OPEN SHEET
' ============================================================
Public Sub OpenSheetFromBtn()

    Dim shp As Shape

    Set shp = ActiveSheet.Shapes(Application.Caller)

    On Error Resume Next

    ActiveWorkbook.Sheets(shp.AlternativeText).Activate

    On Error GoTo 0

End Sub

' ============================================================
' DELETE SHEET FROM INDEX
' ============================================================
Public Sub DeleteSheetFromIndex()

    On Error GoTo errHandle

    Dim btnName As String
    btnName = Application.Caller

    Dim btnNo As Integer
    btnNo = CInt(Replace(btnName, "DelBtn_", ""))

    Dim idx As Worksheet
    Set idx = ActiveWorkbook.Sheets("Sheet_Index")

    Dim r As Long
    Dim sheetName As String

    For r = 4 To idx.UsedRange.Rows.count + 3

        If idx.Cells(r, 1).Value = btnNo Then

            sheetName = idx.Cells(r, 7).Value

            Exit For

        End If

    Next r

    If sheetName = "" Then Exit Sub

    Dim ask As Integer

    ask = MsgBox( _
        "Delete sheet '" & sheetName & "' ?", _
        vbYesNo + vbCritical)

    If ask = vbNo Then Exit Sub

    Application.DisplayAlerts = False

    ActiveWorkbook.Sheets(sheetName).Delete

    Application.DisplayAlerts = True

    Dim dummy As IRibbonControl

    Call GenerateSheetIndex(dummy)

    Exit Sub

errHandle:

    MsgBox Err.Description, vbExclamation

End Sub
' ============================================================
' ADD BACK BUTTONS — bottom-right, uses Forms Button (no TextFrame2)
' ============================================================
Public Sub AddBackButtons(control As IRibbonControl)

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim answer As Integer
    answer = MsgBox("Back button goes to:" & vbNewLine & _
                    "Yes = Sheet_Gallery" & vbNewLine & _
                    "No  = Sheet_Index", _
                    vbYesNo + vbQuestion, "XLSmart - Back Button Target")

    Dim targetSheet As String
    If answer = vbYes Then
        targetSheet = "Sheet_Gallery"
    Else
        targetSheet = "Sheet_Index"
    End If

    ' Store target
    On Error Resume Next
    wb.Names("XLS_BackTarget").Delete
    On Error GoTo 0
    wb.Names.Add Name:="XLS_BackTarget", _
                 RefersTo:="=""" & targetSheet & """"

    Dim ws As Worksheet
    For Each ws In wb.Sheets
        If ws.Name = "Sheet_Index" Or ws.Name = "Sheet_Gallery" Then GoTo SkipBack

        ' Remove old back button
        Dim shp As Shape
        For Each shp In ws.Shapes
            If shp.Name = "XLS_BackBtn" Then
                shp.Delete
                Exit For
            End If
        Next shp

        ' Also remove any Forms buttons named XLS_BackBtn
        Dim ob As Object
        For Each ob In ws.Buttons
            If ob.Name = "XLS_BackBtn" Then
                ob.Delete
                Exit For
            End If
        Next ob

        ' Find bottom-right position
        Dim lastR As Long, lastC As Long
        On Error Resume Next
        lastR = ws.UsedRange.Rows(ws.UsedRange.Rows.count).Row
        lastC = ws.UsedRange.Columns(ws.UsedRange.Columns.count).Column
        On Error GoTo 0
        If lastR < 1 Then lastR = 1
        If lastC < 1 Then lastC = 1

        ' Button dimensions
        Dim btnL As Double, btnT As Double
        Dim btnW As Double, btnH As Double
        btnW = 110
        btnH = 22

        ' Position: top-left (A1)
        btnL = ws.Cells(1, 1).left + 5
        btnT = ws.Cells(1, 1).Top + 5

        ' Use Forms Button   works on ALL Excel versions, no TextFrame2 needed
        Dim Btn As Object
        Set Btn = ws.Buttons.Add(btnL, btnT, btnW, btnH)
        Btn.Name = "XLS_BackBtn"
        Btn.Caption = "<< Back to Home"
        Btn.Font.Size = 9
        Btn.Font.Bold = True
        Btn.OnAction = "GoBackHome"

SkipBack:
    Next ws

    MsgBox "Back to Home buttons added to all sheets!", vbInformation, "XLSmart"
End Sub

' ============================================================
' GO BACK HOME
' ============================================================
Public Sub GoBackHome()
    Dim targetSheet As String
    On Error Resume Next
    targetSheet = Evaluate(ActiveWorkbook.Names("XLS_BackTarget").RefersTo)
    On Error GoTo 0

    If Trim(targetSheet) = "" Then targetSheet = "Sheet_Index"

    On Error Resume Next
    ActiveWorkbook.Sheets(targetSheet).Activate
    If Err.Number <> 0 Then
        MsgBox "Home sheet not found. Please generate Sheet Index or Gallery first.", _
               vbExclamation, "XLSmart"
    End If
    On Error GoTo 0
End Sub

' ============================================================
' REMOVE BACK BUTTONS
' ============================================================
Public Sub RemoveBackButtons(control As IRibbonControl)
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim ws As Worksheet
    Dim ob As Object
    For Each ws In wb.Sheets
        ' Remove Forms buttons
        For Each ob In ws.Buttons
            If ob.Name = "XLS_BackBtn" Then ob.Delete
        Next ob
        ' Remove any Shape-based buttons
        Dim shp As Shape
        For Each shp In ws.Shapes
            If shp.Name = "XLS_BackBtn" Then shp.Delete
        Next shp
    Next ws

    MsgBox "All Back to Home buttons removed!", vbInformation, "XLSmart"
End Sub


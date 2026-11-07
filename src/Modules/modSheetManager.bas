Attribute VB_Name = "modSheetManager"
Option Explicit

' ============================================================
' MODULE: modSheetManager
' XLSmart - Control Group - Sheet Manager
' ============================================================

Public Sub SheetManager_Main(control As IRibbonControl)
    MsgBox "Sheet Manager - Select a specific option from the dropdown menu.", _
           vbInformation, "XLSmart - Sheet Manager"
End Sub

' ----- UPPERCASE -----
Public Sub SheetManager_Uppercase(control As IRibbonControl)
    Dim cell As Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first!", vbExclamation, "XLSmart"
        Exit Sub
    End If
    For Each cell In Selection
        If Not cell.HasFormula Then
            If Trim(cell.Value) <> "" Then
                cell.Value = UCase(cell.Value)
            End If
        End If
    Next cell
    MsgBox "Text converted to UPPERCASE successfully!", vbInformation, "XLSmart"
End Sub

' ----- lowercase -----
Public Sub SheetManager_Lowercase(control As IRibbonControl)
    Dim cell As Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first!", vbExclamation, "XLSmart"
        Exit Sub
    End If
    For Each cell In Selection
        If Not cell.HasFormula Then
            If Trim(cell.Value) <> "" Then
                cell.Value = LCase(cell.Value)
            End If
        End If
    Next cell
    MsgBox "Text converted to lowercase successfully!", vbInformation, "XLSmart"
End Sub

' ----- Proper Case -----
Public Sub SheetManager_ProperCase(control As IRibbonControl)
    Dim cell As Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first!", vbExclamation, "XLSmart"
        Exit Sub
    End If
    For Each cell In Selection
        If Not cell.HasFormula Then
            If Trim(cell.Value) <> "" Then
                cell.Value = WorksheetFunction.Proper(cell.Value)
            End If
        End If
    Next cell
    MsgBox "Text converted to Proper Case successfully!", vbInformation, "XLSmart"
End Sub

' ----- Import Sheets -----
Public Sub SheetManager_ImportSheets(control As IRibbonControl)

    Dim fd As FileDialog

    Dim srcWB As Workbook
    Dim destWB As Workbook

    Dim ws As Worksheet
    Dim copiedWs As Worksheet

    Dim filePath As String
    Dim newName As String

    Dim okCount As Long
    Dim skipCount As Long

    Dim logText As String

    Dim hasHidden As Boolean
    Dim importHidden As VbMsgBoxResult

    Dim oldState As XlSheetVisibility


    Set destWB = ActiveWorkbook

    If destWB Is Nothing Then

        MsgBox "No active workbook found.", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If


    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd

        .Title = "Select Source Workbook"

        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb"

        .AllowMultiSelect = False

        If .Show <> -1 Then Exit Sub

        filePath = .SelectedItems(1)

    End With


    On Error GoTo MainErr

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Set srcWB = Workbooks.Open(filePath, ReadOnly:=True)


    ' workbook structure protection
    If srcWB.ProtectStructure Then

        MsgBox "Source workbook structure is protected." & vbCrLf & _
               "Please unprotect workbook first.", _
               vbExclamation, "XLSmart"

        GoTo SafeExit

    End If


    ' check hidden sheets
    For Each ws In srcWB.Worksheets

        If ws.Visible <> xlSheetVisible Then

            hasHidden = True
            Exit For

        End If

    Next ws


    ' ask user
    If hasHidden Then

        importHidden = MsgBox( _
            "This workbook contains hidden sheets." & vbCrLf & vbCrLf & _
            "YES  = Import all sheets including hidden sheets" & vbCrLf & _
            "NO   = Skip hidden sheets", _
            vbYesNo + vbQuestion, _
            "XLSmart - Hidden Sheets Found")

    End If



    ' import sheets
    For Each ws In srcWB.Worksheets

        On Error GoTo SheetErr


        ' skip hidden sheets if user clicked NO
        If hasHidden Then

            If importHidden = vbNo Then

                If ws.Visible <> xlSheetVisible Then

                    skipCount = skipCount + 1

                    logText = logText & vbCrLf & _
                              "Skipped Hidden: " & ws.Name

                    GoTo NextSheet

                End If

            End If

        End If


        newName = ws.Name

        If SheetExists(newName, destWB) Then
            newName = MakeUniqueSheetName(newName, destWB)
        End If


        ' save original hidden state
        oldState = ws.Visible


        ' temporary visible for safe copy
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
        End If


        ' safer copy
        ws.Copy

        Set copiedWs = ActiveWorkbook.Sheets(1)

        copiedWs.Move After:=destWB.Sheets(destWB.Sheets.count)



        ' rename copied sheet
        Err.Clear

        On Error Resume Next

        copiedWs.Name = newName

        If Err.Number <> 0 Then

            copiedWs.Name = "Sheet_" & _
                            Format(Timer * 100, "0")

            Err.Clear

        End If

        On Error GoTo SheetErr



        ' restore hidden state in imported sheet
        Err.Clear

        On Error Resume Next

        copiedWs.Visible = oldState
        Err.Clear

        ws.Visible = oldState
        Err.Clear

        On Error GoTo SheetErr



        ' successful import
        okCount = okCount + 1


NextSheet:

        Err.Clear
        On Error GoTo 0

    Next ws



    MsgBox okCount & " sheet(s) imported successfully." & _
           vbCrLf & skipCount & " skipped." & _
           vbCrLf & logText, _
           vbInformation, "XLSmart"

    GoTo SafeExit



' per sheet error
SheetErr:

    If Err.Number <> 0 Then

        skipCount = skipCount + 1

        logText = logText & vbCrLf & _
                  "Failed: " & ws.Name & _
                  " | " & Err.Description

    End If

    Err.Clear
    Resume Next



MainErr:

    MsgBox "Import failed : " & Err.Description, _
           vbCritical, "XLSmart"



SafeExit:

    On Error Resume Next

    If Not srcWB Is Nothing Then
        srcWB.Close SaveChanges:=False
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub




' ----- Check Existing Sheet -----
Private Function SheetExists(sName As String, wb As Workbook) As Boolean

    Dim sht As Worksheet

    For Each sht In wb.Worksheets

        If LCase(Trim(sht.Name)) = LCase(Trim(sName)) Then

            SheetExists = True
            Exit Function

        End If

    Next sht

End Function




' ----- Create Unique Sheet Name -----
Private Function MakeUniqueSheetName(baseName As String, wb As Workbook) As String

    Dim tempName As String
    Dim n As Long

    tempName = left(baseName, 25)

    n = 1

    Do While SheetExists(tempName & "_" & n, wb)

        n = n + 1

    Loop

    MakeUniqueSheetName = tempName & "_" & n

End Function

' ----- Clean Sheet -----
Public Sub SheetManager_CleanSheet(control As IRibbonControl)
    Dim ans As Integer
    ans = MsgBox("Are you sure you want to clear all data from the current sheet?", _
                 vbYesNo + vbQuestion, "XLSmart - Clean Sheet")
    If ans = vbNo Then Exit Sub
    
    On Error GoTo ErrHandler
    ActiveSheet.Cells.Clear
    MsgBox "Sheet cleaned successfully!", vbInformation, "XLSmart"
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "XLSmart"
End Sub

' ----- Auto-Fit Sheets -----
Public Sub SheetManager_AutoFitSheets(control As IRibbonControl)

    Dim wb As Workbook
    Dim ws As Worksheet

    Dim lastRow As Long
    Dim lastCol As Long

    Dim rng As Range

    Dim okCount As Long
    Dim skipCount As Long

    Dim logText As String

    Dim ans As VbMsgBoxResult
    Dim onlyActive As Boolean


    Set wb = ActiveWorkbook

    If wb Is Nothing Then

        MsgBox "No active workbook found.", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If


    ' ask user
    ans = MsgBox( _
        "YES = Apply Auto-Fit to Active Sheet only" & vbCrLf & _
        "NO = Apply Auto-Fit to All Sheets", _
        vbYesNo + vbQuestion, _
        "XLSmart - AutoFit")



    If ans = vbYes Then
        onlyActive = True
    Else
        onlyActive = False
    End If



    Application.ScreenUpdating = False



    ' ----- ACTIVE SHEET ONLY -----
    If onlyActive Then

        Set ws = ActiveSheet

        On Error GoTo SheetErr


        If ws.ProtectContents Then

            MsgBox "Active sheet is protected.", _
                   vbExclamation, "XLSmart"

            GoTo SafeExit

        End If


        lastRow = 0
        lastCol = 0

        On Error Resume Next

        lastRow = ws.Cells.Find(What:="*", _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious).Row

        lastCol = ws.Cells.Find(What:="*", _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlPrevious).Column

        On Error GoTo SheetErr


        If lastRow > 0 And lastCol > 0 Then

            Set rng = ws.Range(ws.Cells(1, 1), _
                               ws.Cells(lastRow, lastCol))

            rng.Columns.AutoFit

            On Error Resume Next
            rng.Rows.AutoFit
            Err.Clear
            On Error GoTo SheetErr

        End If


        Application.ScreenUpdating = True

        MsgBox "Auto-Fit applied to active sheet successfully!", _
               vbInformation, "XLSmart"

        Exit Sub

    End If




    ' ----- ALL SHEETS -----
    For Each ws In wb.Worksheets

        On Error GoTo SheetErr


        If ws.ProtectContents Then

            skipCount = skipCount + 1

            logText = logText & vbCrLf & _
                      "Skipped Protected: " & ws.Name

            GoTo NextSheet

        End If


        lastRow = 0
        lastCol = 0

        On Error Resume Next

        lastRow = ws.Cells.Find(What:="*", _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious).Row

        lastCol = ws.Cells.Find(What:="*", _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlPrevious).Column

        On Error GoTo SheetErr


        If lastRow = 0 Or lastCol = 0 Then
            GoTo NextSheet
        End If


        Set rng = ws.Range(ws.Cells(1, 1), _
                           ws.Cells(lastRow, lastCol))

        rng.Columns.AutoFit

        On Error Resume Next
        rng.Rows.AutoFit
        Err.Clear
        On Error GoTo SheetErr


        okCount = okCount + 1


NextSheet:

        Err.Clear
        On Error GoTo 0

    Next ws


    Application.ScreenUpdating = True


    MsgBox "Auto-Fit applied successfully!" & _
           vbCrLf & vbCrLf & _
           "Success : " & okCount & vbCrLf & _
           "Skipped : " & skipCount & _
           vbCrLf & logText, _
           vbInformation, "XLSmart"

    Exit Sub



SheetErr:

    skipCount = skipCount + 1

    logText = logText & vbCrLf & _
              "Failed: " & ws.Name & _
              " | " & Err.Description

    Err.Clear
    Resume Next



SafeExit:

    Application.ScreenUpdating = True

End Sub

' ----- Hide Sheets -----
Public Sub SheetManager_HideSheets(control As IRibbonControl)

    Dim wb As Workbook
    Dim ws As Worksheet

    Dim activeWS As Worksheet

    Dim hideCount As Long
    Dim skipCount As Long

    Dim logText As String
    Dim finalMsg As String


    Set wb = ActiveWorkbook

    If wb Is Nothing Then

        MsgBox "No active workbook found.", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If


    Set activeWS = ActiveSheet


    On Error GoTo ErrHandler

    Application.ScreenUpdating = False


    For Each ws In wb.Worksheets

        If ws.Name <> activeWS.Name Then

            On Error Resume Next

            ws.Visible = xlSheetHidden

            If Err.Number = 0 Then

                hideCount = hideCount + 1

            Else

                skipCount = skipCount + 1

                logText = logText & vbCrLf & _
                          "Failed: " & ws.Name

                Err.Clear

            End If

            On Error GoTo ErrHandler

        End If

    Next ws


    Application.ScreenUpdating = True


    finalMsg = hideCount & " sheet(s) hidden successfully!" & _
               vbCrLf & vbCrLf & _
               "Active Sheet : " & activeWS.Name


    If skipCount > 0 Then

        finalMsg = finalMsg & _
                   vbCrLf & vbCrLf & _
                   "Skipped : " & skipCount & _
                   vbCrLf & logText

    End If


    MsgBox finalMsg, vbInformation, "XLSmart"

    Exit Sub


ErrHandler:

    Application.ScreenUpdating = True

    MsgBox "Error: " & Err.Description, _
           vbCritical, "XLSmart"

End Sub





' ----- Unhide Sheets -----
Public Sub SheetManager_UnhideSheets(control As IRibbonControl)

    Dim wb As Workbook
    Dim ws As Worksheet

    Dim count As Long
    Dim skipCount As Long

    Dim logText As String
    Dim finalMsg As String


    Set wb = ActiveWorkbook

    If wb Is Nothing Then

        MsgBox "No active workbook found.", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If


    On Error GoTo ErrHandler

    Application.ScreenUpdating = False


    For Each ws In wb.Worksheets

        If ws.Visible <> xlSheetVisible Then

            On Error Resume Next

            ws.Visible = xlSheetVisible

            If Err.Number = 0 Then

                count = count + 1

            Else

                skipCount = skipCount + 1

                logText = logText & vbCrLf & _
                          "Failed: " & ws.Name

                Err.Clear

            End If

            On Error GoTo ErrHandler

        End If

    Next ws


    Application.ScreenUpdating = True


    finalMsg = count & " sheet(s) unhidden successfully!"


    If skipCount > 0 Then

        finalMsg = finalMsg & _
                   vbCrLf & vbCrLf & _
                   "Skipped : " & skipCount & _
                   vbCrLf & logText

    End If


    MsgBox finalMsg, vbInformation, "XLSmart"

    Exit Sub


ErrHandler:

    Application.ScreenUpdating = True

    MsgBox "Error: " & Err.Description, _
           vbCritical, "XLSmart"

End Sub

' ----- Hide Shapes -----
Public Sub SheetManager_HideShapes(control As IRibbonControl)
    Dim shp As Shape
    Dim count As Integer
    count = 0
    
    On Error GoTo ErrHandler
    For Each shp In ActiveSheet.Shapes
        shp.Visible = False
        count = count + 1
    Next shp
    MsgBox count & " shape(s) hidden successfully!", vbInformation, "XLSmart"
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "XLSmart"
End Sub

' ----- Unhide Shapes -----
Public Sub SheetManager_UnhideShapes(control As IRibbonControl)
    Dim shp As Shape
    Dim count As Integer
    count = 0
    
    On Error GoTo ErrHandler
    For Each shp In ActiveSheet.Shapes
        shp.Visible = True
        count = count + 1
    Next shp
    MsgBox count & " shape(s) unhidden successfully!", vbInformation, "XLSmart"
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "XLSmart"
End Sub

' ----- Select / Unselect Shapes -----
Public Sub SheetManager_SelectShapes(control As IRibbonControl)

    Dim ws As Worksheet
    Dim shp As Shape

    Dim count As Long

    Static shapesSelected As Boolean


    Set ws = ActiveSheet


    On Error GoTo ErrHandler


    If ws.Shapes.count = 0 Then

        MsgBox "No shapes found on the current sheet!", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If



    ' ----- UNSELECT -----
    If shapesSelected = True Then

        ws.Range("A1").Select

        shapesSelected = False

        MsgBox "Shapes unselected successfully!", _
               vbInformation, "XLSmart"

        Exit Sub

    End If



    ' ----- SELECT -----
    For Each shp In ws.Shapes
        count = count + 1
    Next shp


    ws.Shapes.SelectAll

    shapesSelected = True


    MsgBox count & " shape(s) selected!", _
           vbInformation, "XLSmart"

    Exit Sub



ErrHandler:

    MsgBox "Error: " & Err.Description, _
           vbCritical, "XLSmart"

End Sub

' ----- Values to Text Format -----
Public Sub SheetManager_ValuesToText(control As IRibbonControl)
    Dim cell As Range
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first!", vbExclamation, "XLSmart"
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    For Each cell In Selection
        If Not cell.HasFormula Then
            Dim val As String
            val = CStr(cell.Value)
            cell.NumberFormat = "@"
            cell.Value = val
        End If
    Next cell
    Application.ScreenUpdating = True
    MsgBox "Values converted to Text format successfully!", vbInformation, "XLSmart"
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical, "XLSmart"
End Sub

' ----- Duplicate Active Sheet -----
Public Sub SheetManager_DuplicateSheet(control As IRibbonControl)

    Dim wb As Workbook
    Dim srcSheet As Worksheet
    Dim newSheet As Worksheet

    Dim baseName As String
    Dim tryName As String

    Dim n As Long
    Dim existsOk As Boolean

    Dim ws As Worksheet


    Set wb = ActiveWorkbook

    If wb Is Nothing Then

        MsgBox "No active workbook found.", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If


    Set srcSheet = ActiveSheet


    On Error GoTo ErrHandler

    Application.ScreenUpdating = False


    ' duplicate sheet
    srcSheet.Copy After:=wb.Sheets(wb.Sheets.count)

    Set newSheet = wb.Sheets(wb.Sheets.count)



    ' name style -> SheetName (1)
    baseName = srcSheet.Name

    n = 1

CheckAgain:

    tryName = baseName & " (" & n & ")"

    existsOk = False


    For Each ws In wb.Worksheets

        If ws.Name = tryName And ws.Index <> newSheet.Index Then

            existsOk = True
            Exit For

        End If

    Next ws


    If existsOk = True Then

        n = n + 1
        GoTo CheckAgain

    End If


    newSheet.Name = tryName


    Application.ScreenUpdating = True


    MsgBox "Sheet duplicated successfully!" & _
           vbCrLf & vbCrLf & _
           "New Sheet : " & newSheet.Name, _
           vbInformation, "XLSmart"

    Exit Sub



ErrHandler:

    Application.ScreenUpdating = True

    MsgBox "Error: " & Err.Description, _
           vbCritical, "XLSmart"

End Sub

' ----- Sort Sheets A-Z -----
Public Sub SheetManager_SortAZ(control As IRibbonControl)

    SortSheets True

End Sub




' ----- Sort Sheets Z-A -----
Public Sub SheetManager_SortZA(control As IRibbonControl)

    SortSheets False

End Sub




' ----- Main Sort Function -----
Private Sub SortSheets(isAZ As Boolean)

    Dim wb As Workbook

    Dim i As Long
    Dim j As Long

    Dim moved As Boolean

    Dim msgText As String


    Set wb = ActiveWorkbook

    If wb Is Nothing Then

        MsgBox "No active workbook found.", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If


    On Error GoTo ErrHandler

    Application.ScreenUpdating = False


    ' bubble sort style
    Do

        moved = False

        For i = 1 To wb.Worksheets.count - 1

            If isAZ = True Then

                If UCase(wb.Worksheets(i).Name) > _
                   UCase(wb.Worksheets(i + 1).Name) Then

                    wb.Worksheets(i + 1).Move _
                    Before:=wb.Worksheets(i)

                    moved = True

                End If

            Else

                If UCase(wb.Worksheets(i).Name) < _
                   UCase(wb.Worksheets(i + 1).Name) Then

                    wb.Worksheets(i + 1).Move _
                    Before:=wb.Worksheets(i)

                    moved = True

                End If

            End If

        Next i

    Loop While moved = True


    Application.ScreenUpdating = True


    If isAZ = True Then

        msgText = "Sheets sorted A-Z successfully!"

    Else

        msgText = "Sheets sorted Z-A successfully!"

    End If


    MsgBox msgText, vbInformation, "XLSmart"

    Exit Sub



ErrHandler:

    Application.ScreenUpdating = True

    MsgBox "Error: " & Err.Description, _
           vbCritical, "XLSmart"

End Sub

' ----- Create Monthly Sheets -----
Public Sub SheetManager_MonthlySheets(control As IRibbonControl)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim s As Worksheet

    Dim i As Long

    Dim fullMonths As Variant
    Dim shortMonths As Variant

    Dim shName As String

    Dim existsOk As Boolean

    Dim createdCount As Long
    Dim skipCount As Long

    Dim ans As VbMsgBoxResult

    Dim yearTxt As String

    Dim finalMsg As String
    Dim logText As String


    Set wb = ActiveWorkbook

    If wb Is Nothing Then

        MsgBox "No active workbook found.", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If


    ' month arrays
    fullMonths = Array( _
        "January", "February", "March", "April", _
        "May", "June", "July", "August", _
        "September", "October", "November", "December")

    shortMonths = Array( _
        "Jan", "Feb", "Mar", "Apr", _
        "May", "Jun", "Jul", "Aug", _
        "Sep", "Oct", "Nov", "Dec")



    ' ask user
    ans = MsgBox( _
        "YES = Create full month sheets" & vbCrLf & _
        "Example : January, February" & vbCrLf & vbCrLf & _
        "NO = Create short month with current year" & vbCrLf & _
        "Example : Jan-" & Format(Date, "yy"), _
        vbYesNo + vbQuestion, _
        "XLSmart - Monthly Sheets")



    yearTxt = Format(Date, "yy")


    On Error GoTo ErrHandler

    Application.ScreenUpdating = False



    For i = 0 To 11

        ' YES = January
        If ans = vbYes Then

            shName = fullMonths(i)

        Else

            ' NO = Jan-26
            shName = shortMonths(i) & "-" & yearTxt

        End If



        ' check existing sheet
        existsOk = False

        For Each s In wb.Worksheets

            If LCase(Trim(s.Name)) = _
               LCase(Trim(shName)) Then

                existsOk = True
                Exit For

            End If

        Next s



        ' create sheet
        If existsOk = False Then

            Set ws = wb.Worksheets.Add( _
                     After:=wb.Sheets(wb.Sheets.count))

            ws.Name = shName

            createdCount = createdCount + 1

        Else

            skipCount = skipCount + 1

            logText = logText & vbCrLf & _
                      "Already Exists: " & shName

        End If

    Next i



    Application.ScreenUpdating = True



    finalMsg = createdCount & _
               " monthly sheet(s) created successfully!"


    If skipCount > 0 Then

        finalMsg = finalMsg & _
                   vbCrLf & vbCrLf & _
                   "Skipped : " & skipCount & _
                   vbCrLf & logText

    End If


    MsgBox finalMsg, vbInformation, "XLSmart"

    Exit Sub



ErrHandler:

    Application.ScreenUpdating = True

    MsgBox "Error creating monthly sheets: " & _
           Err.Description, _
           vbCritical, "XLSmart"

End Sub

' ----- Remove Blank Sheets -----
Public Sub SheetManager_RemoveBlankSheets(control As IRibbonControl)

    Dim wb As Workbook
    Dim ws As Worksheet

    Dim wsNames() As String

    Dim delCount As Long
    Dim removedCount As Long

    Dim k As Long

    Dim ans As VbMsgBoxResult

    Dim finalMsg As String


    Set wb = ActiveWorkbook

    If wb Is Nothing Then

        MsgBox "No active workbook found.", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If



    ans = MsgBox( _
        "This will permanently delete all blank sheets" & vbCrLf & _
        "from the active workbook." & vbCrLf & vbCrLf & _
        "Continue?", _
        vbYesNo + vbQuestion, _
        "XLSmart - Remove Blank Sheets")


    If ans = vbNo Then Exit Sub



    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False



    ReDim wsNames(1 To wb.Worksheets.count)



    ' collect blank sheet names
    For Each ws In wb.Worksheets

        ' keep at least 1 sheet
        If wb.Worksheets.count - removedCount <= 1 Then
            Exit For
        End If


        If WorksheetFunction.CountA(ws.Cells) = 0 Then

            delCount = delCount + 1

            wsNames(delCount) = ws.Name

        End If

    Next ws




    ' delete sheets
    For k = 1 To delCount

        On Error Resume Next

        wb.Worksheets(wsNames(k)).Delete

        If Err.Number = 0 Then
            removedCount = removedCount + 1
        End If

        Err.Clear

        On Error GoTo ErrHandler

    Next k



    Application.DisplayAlerts = True
    Application.ScreenUpdating = True



    If removedCount = 0 Then

        MsgBox "No blank sheets found.", _
               vbInformation, "XLSmart"

    Else

        finalMsg = removedCount & _
                   " blank sheet(s) removed successfully!"

        MsgBox finalMsg, _
               vbInformation, "XLSmart"

    End If


    Exit Sub



ErrHandler:

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Error: " & Err.Description, _
           vbCritical, "XLSmart"

End Sub

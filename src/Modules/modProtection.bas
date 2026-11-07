Attribute VB_Name = "modProtection"
Option Explicit

' ============================================================
' MODULE: modProtection
' XLBooster - Control Group - Protection
' ============================================================

Public Sub Protection_Main(control As IRibbonControl)
    MsgBox "Protection - Select a specific option from the dropdown menu.", _
           vbInformation, "XLBooster - Protection"
End Sub

' ----- Backup File -----
Public Sub Protection_BackupFile(control As IRibbonControl)

    Dim wb As Workbook

    Dim fd As FileDialog

    Dim defaultName As String
    Dim backupName As String

    Dim savePath As String

    Dim ans As VbMsgBoxResult

    Dim fileExt As String
    Dim baseName As String

    Dim timestamp As String


    Set wb = ActiveWorkbook

    If wb Is Nothing Then

        MsgBox "No active workbook found.", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If



    ' workbook must be saved first
    If wb.Path = "" Then

        MsgBox "Please save the workbook first before creating a backup!", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If



    ' create dynamic backup name
    timestamp = Format(Now, "yyyy-mm-dd_hh-nn-ss")

    baseName = left(wb.Name, InStrRev(wb.Name, ".") - 1)

    fileExt = Mid(wb.Name, InStrRev(wb.Name, "."))

    backupName = baseName & "_Backup_" & timestamp & fileExt



    ' custom popup
    ans = MsgBox( _
        "BACKUP WORKBOOK BEFORE PROTECTION" & vbCrLf & _
        "===================================" & vbCrLf & vbCrLf & _
        "It is recommended to create a backup." & vbCrLf & vbCrLf & _
        "New Name: " & backupName & vbCrLf & _
        "Do you want to create a backup now?", _
        vbYesNo + vbQuestion, _
        "Create Backup?")


    If ans = vbNo Then Exit Sub



    ' save dialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)

    With fd

        .Title = "Save Backup File"

        .InitialFileName = wb.Path & "\" & backupName

        If .Show <> -1 Then Exit Sub

        savePath = .SelectedItems(1)

    End With



    On Error GoTo ErrHandler


    Application.ScreenUpdating = False


    ' create backup copy
    wb.SaveCopyAs savePath


    Application.ScreenUpdating = True


    MsgBox "Backup saved successfully!" & _
           vbCrLf & vbCrLf & _
           savePath, _
           vbInformation, "XLSmart"

    Exit Sub



ErrHandler:

    Application.ScreenUpdating = True

    MsgBox "Error creating backup: " & _
           Err.Description, _
           vbCritical, "XLSmart"

End Sub

' ----- Protect Workbook Sheets -----
Public Sub Protection_ProtectSheets(control As IRibbonControl)

    Dim wb As Workbook
    Dim ws As Worksheet

    Dim passWord As String

    Dim ans As VbMsgBoxResult

    Dim protectedCount As Long
    Dim skipCount As Long

    Dim logText As String
    Dim finalMsg As String

    Dim isStrict As Boolean


    Set wb = ActiveWorkbook

    If wb Is Nothing Then

        MsgBox "No active workbook found.", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If



    ' password input
    passWord = InputBox( _
        "Enter protection password", _
        "XLSmart - Sheet Protection")


    If Trim(passWord) = "" Then Exit Sub



    ' protection level popup
    ans = MsgBox( _
        "Select protection level" & vbCrLf & vbCrLf & _
        "YES = Strict Protection" & vbCrLf & _
        "(no editing of protected cells)" & vbCrLf & vbCrLf & _
        "NO = Flexible Protection" & vbCrLf & _
        "(allow formatting)", _
        vbYesNo + vbQuestion, _
        "Protection Level")



    If ans = vbYes Then

        isStrict = True

    Else

        isStrict = False

    End If



    On Error GoTo ErrHandler

    Application.ScreenUpdating = False



    ' apply protection to all sheets
    For Each ws In wb.Worksheets

        On Error Resume Next


        ws.Unprotect passWord


        If isStrict = True Then

            ' strict protection
            ws.Protect passWord:=passWord, _
                       DrawingObjects:=True, _
                       Contents:=True, _
                       Scenarios:=True

        Else

            ' flexible protection
            ws.Protect passWord:=passWord, _
                       DrawingObjects:=True, _
                       Contents:=True, _
                       Scenarios:=True, _
                       AllowFormattingCells:=True, _
                       AllowFormattingColumns:=True, _
                       AllowFormattingRows:=True

        End If



        If Err.Number = 0 Then

            protectedCount = protectedCount + 1

        Else

            skipCount = skipCount + 1

            logText = logText & vbCrLf & _
                      "Failed: " & ws.Name

            Err.Clear

        End If


        On Error GoTo ErrHandler

    Next ws



    Application.ScreenUpdating = True



    finalMsg = protectedCount & _
               " sheet(s) protected successfully!"


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

' ----- Unprotect All Sheets -----
Public Sub Protection_UnprotectSheets(control As IRibbonControl)

    Dim wb As Workbook
    Dim ws As Worksheet

    Dim pwd As String

    Dim unprotectCount As Long
    Dim skipCount As Long

    Dim logText As String
    Dim finalMsg As String

    Dim wasProtected As Boolean


    Set wb = ActiveWorkbook

    If wb Is Nothing Then

        MsgBox "No active workbook found.", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If



    ' password input
    pwd = InputBox( _
        "Enter password to unprotect all sheets:", _
        "XLSmart - Unprotect Sheets")


    If Trim(pwd) = "" Then

        MsgBox "Unprotect cancelled.", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If



    On Error GoTo ErrHandler

    Application.ScreenUpdating = False



    ' unprotect all sheets
    For Each ws In wb.Worksheets

        wasProtected = ws.ProtectContents Or _
                       ws.ProtectDrawingObjects Or _
                       ws.ProtectScenarios


        If wasProtected Then

            On Error Resume Next

            ws.Unprotect passWord:=pwd


            If ws.ProtectContents = False Then

                unprotectCount = unprotectCount + 1

            Else

                skipCount = skipCount + 1

                logText = logText & vbCrLf & _
                          "Wrong Password: " & ws.Name

            End If


            Err.Clear

            On Error GoTo ErrHandler

        End If

    Next ws



    Application.ScreenUpdating = True



    If unprotectCount = 0 Then

        MsgBox "Incorrect password or no protected sheets found!", _
               vbExclamation, "XLSmart"

    Else

        finalMsg = unprotectCount & _
                   " sheet(s) unprotected successfully!"


        If skipCount > 0 Then

            finalMsg = finalMsg & _
                       vbCrLf & vbCrLf & _
                       "Skipped : " & skipCount & _
                       vbCrLf & logText

        End If


        MsgBox finalMsg, _
               vbInformation, "XLSmart"

    End If


    Exit Sub



ErrHandler:

    Application.ScreenUpdating = True

    MsgBox "Error unprotecting sheets: " & _
           Err.Description, _
           vbCritical, "XLSmart"

End Sub

' ----- Protect Workbook Structure -----
Public Sub Protection_ProtectWorkbook(control As IRibbonControl)

    Dim wb As Workbook

    Dim pwd As String
    Dim confirmPwd As String

    Dim ans As VbMsgBoxResult


    Set wb = ActiveWorkbook

    If wb Is Nothing Then

        MsgBox "No active workbook found.", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If



    ' password input
    pwd = InputBox( _
        "Enter password to protect workbook structure:" & vbCrLf & vbCrLf & _
        "Workbook data will still be editable." & vbCrLf & _
        "But sheets cannot be:" & vbCrLf & _
        "- Deleted" & vbCrLf & _
        "- Renamed" & vbCrLf & _
        "- Hidden / Unhidden" & vbCrLf & _
        "- Added / Moved", _
        "XLSmart - Protect Workbook")


    If Trim(pwd) = "" Then Exit Sub



    ' confirm password
    confirmPwd = InputBox( _
        "Re-enter password to confirm:", _
        "XLSmart - Confirm Password")


    If pwd <> confirmPwd Then

        MsgBox "Passwords do not match!" & vbCrLf & _
               "Protection cancelled.", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If



    On Error GoTo ErrHandler

    Application.ScreenUpdating = False



    ' protect workbook structure
    wb.Protect passWord:=pwd, _
               Structure:=True, _
               Windows:=False



    Application.ScreenUpdating = True



    MsgBox "Workbook structure protected successfully!" & _
           vbCrLf & vbCrLf & _
           "Users can still:" & vbCrLf & _
           "- Edit data" & vbCrLf & _
           "- Format cells" & vbCrLf & vbCrLf & _
           "But sheet management options are now disabled.", _
           vbInformation, "XLSmart"

    Exit Sub



ErrHandler:

    Application.ScreenUpdating = True

    MsgBox "Error protecting workbook: " & _
           Err.Description, _
           vbCritical, "XLSmart"

End Sub





' ----- Unprotect Workbook Structure -----
Public Sub Protection_UnprotectWorkbook(control As IRibbonControl)

    Dim wb As Workbook

    Dim pwd As String


    Set wb = ActiveWorkbook

    If wb Is Nothing Then

        MsgBox "No active workbook found.", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If



    ' password input
    pwd = InputBox( _
        "Enter password to unprotect workbook:", _
        "XLSmart - Unprotect Workbook")


    If Trim(pwd) = "" Then Exit Sub



    On Error GoTo ErrHandler

    Application.ScreenUpdating = False



    ' unprotect workbook
    wb.Unprotect passWord:=pwd



    Application.ScreenUpdating = True



    MsgBox "Workbook unprotected successfully!" & _
           vbCrLf & vbCrLf & _
           "Sheet options are now enabled again.", _
           vbInformation, "XLSmart"

    Exit Sub



ErrHandler:

    Application.ScreenUpdating = True

    MsgBox "Incorrect password or workbook is not protected!", _
           vbExclamation, "XLSmart"

End Sub

' ----- Lock Formulas -----
Public Sub Protection_LockFormulas(control As IRibbonControl)
    Dim ws As Worksheet
    Dim pwd As String
    Dim confirmPwd As String
    
    Set ws = ActiveSheet
    
    pwd = InputBox("Enter password to lock all formula cells:", "XLBooster - Lock Formulas")
    If pwd = "" Then
        MsgBox "Lock cancelled. No password entered.", vbExclamation, "XLBooster"
        Exit Sub
    End If
    
    If Len(pwd) < 4 Then
        MsgBox "Password must be at least 4 characters!", vbExclamation, "XLBooster"
        Exit Sub
    End If
    
    confirmPwd = InputBox("Re-enter password to confirm:", "XLBooster - Confirm Password")
    If pwd <> confirmPwd Then
        MsgBox "Passwords do not match! Lock cancelled.", vbExclamation, "XLBooster"
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    ' Unprotect if already protected
    On Error Resume Next
    ws.Unprotect passWord:=pwd
    On Error GoTo ErrHandler
    
    ' Unlock all cells first
    ws.Cells.Locked = False
    
    ' Lock only formula cells
    Dim formulaRange As Range
    On Error Resume Next
    Set formulaRange = ws.Cells.SpecialCells(xlCellTypeFormulas)
    On Error GoTo ErrHandler
    
    If formulaRange Is Nothing Then
        MsgBox "No formula cells found on this sheet!", vbExclamation, "XLBooster"
        Exit Sub
    End If
    
    formulaRange.Locked = True
    ws.Protect passWord:=pwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    MsgBox "All formula cells locked successfully!" & vbNewLine & _
           "Data cells can still be edited.", vbInformation, "XLBooster"
    Exit Sub
ErrHandler:
    MsgBox "Error locking formulas: " & Err.Description, vbCritical, "XLBooster"
End Sub

' ----- Unlock Formulas -----
Public Sub Protection_UnlockFormulas(control As IRibbonControl)
    Dim ws As Worksheet
    Dim pwd As String
    
    Set ws = ActiveSheet
    
    pwd = InputBox("Enter password to unlock formula cells:", "XLBooster - Unlock Formulas")
    If pwd = "" Then
        MsgBox "Unlock cancelled.", vbExclamation, "XLBooster"
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    ws.Unprotect passWord:=pwd
    ws.Cells.Locked = False
    MsgBox "Formula cells unlocked successfully!", vbInformation, "XLBooster"
    Exit Sub
ErrHandler:
    MsgBox "Incorrect password or sheet is not protected!", vbExclamation, "XLBooster"
End Sub

' ----- Lock Selection -----
Public Sub Protection_LockSelection(control As IRibbonControl)

    Dim ws As Worksheet

    Dim pwd As String
    Dim confirmPwd As String

    Dim lockedCount As Double

    Dim ans As VbMsgBoxResult


    If TypeName(Selection) <> "Range" Then

        MsgBox "Please select a range first!", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If


    Set ws = ActiveSheet

    lockedCount = Selection.CountLarge



    ' large selection warning
    If lockedCount > 5000 Then

        ans = MsgBox( _
            "You have selected a very large range." & vbCrLf & _
            "This may take some time." & vbCrLf & vbCrLf & _
            "Continue?", _
            vbYesNo + vbQuestion, _
            "Large Selection")

        If ans = vbNo Then Exit Sub

    End If



    ' password input
    pwd = InputBox( _
        "Enter password to lock selected " & _
        Format(lockedCount, "#,##0") & " cell(s):", _
        "XLSmart - Lock Selection")


    If Trim(pwd) = "" Then

        MsgBox "Lock cancelled. No password entered.", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If



    ' minimum password length
    If Len(pwd) < 4 Then

        MsgBox "Password must be at least 4 characters!", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If



    ' confirm password
    confirmPwd = InputBox( _
        "Re-enter password to confirm:", _
        "XLSmart - Confirm Password")


    If pwd <> confirmPwd Then

        MsgBox "Passwords do not match!" & vbCrLf & _
               "Lock cancelled.", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If



    On Error GoTo ErrHandler

    Application.ScreenUpdating = False



    ' unprotect first
    On Error Resume Next
    ws.Unprotect passWord:=pwd
    Err.Clear
    On Error GoTo ErrHandler



    ' unlock all cells
    ws.Cells.Locked = False


    ' lock only selection
    Selection.Locked = True



    ' protect sheet
    ws.Protect passWord:=pwd, _
               DrawingObjects:=True, _
               Contents:=True, _
               Scenarios:=True



    Application.ScreenUpdating = True



    MsgBox Format(lockedCount, "#,##0") & _
           " cell(s) locked successfully!" & _
           vbCrLf & vbCrLf & _
           "Other cells can still be edited.", _
           vbInformation, "XLSmart"

    Exit Sub



ErrHandler:

    Application.ScreenUpdating = True

    MsgBox "Error locking selection: " & _
           Err.Description, _
           vbCritical, "XLSmart"

End Sub






' ----- Unlock Selection -----
Public Sub Protection_UnlockSelection(control As IRibbonControl)

    Dim ws As Worksheet

    Dim pwd As String

    Dim unlockCount As Double

    Dim ans As VbMsgBoxResult


    If TypeName(Selection) <> "Range" Then

        MsgBox "Please select a range first!", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If


    Set ws = ActiveSheet

    unlockCount = Selection.CountLarge



    ' large selection warning
    If unlockCount > 5000 Then

        ans = MsgBox( _
            "You have selected a very large range." & vbCrLf & _
            "This may take some time." & vbCrLf & vbCrLf & _
            "Continue?", _
            vbYesNo + vbQuestion, _
            "Large Selection")

        If ans = vbNo Then Exit Sub

    End If



    ' password input
    pwd = InputBox( _
        "Enter password to unlock selected cells:", _
        "XLSmart - Unlock Selection")


    If Trim(pwd) = "" Then

        MsgBox "Unlock cancelled.", _
               vbExclamation, "XLSmart"

        Exit Sub

    End If



    On Error GoTo ErrHandler

    Application.ScreenUpdating = False



    ' unprotect sheet
    ws.Unprotect passWord:=pwd



    ' unlock selected cells
    Selection.Locked = False



    ' reprotect sheet
    ws.Protect passWord:=pwd, _
               DrawingObjects:=True, _
               Contents:=True, _
               Scenarios:=True



    Application.ScreenUpdating = True



    MsgBox Format(unlockCount, "#,##0") & _
           " cell(s) unlocked successfully!", _
           vbInformation, "XLSmart"

    Exit Sub



ErrHandler:

    Application.ScreenUpdating = True

    MsgBox "Incorrect password or sheet is not protected!", _
           vbExclamation, "XLSmart"

End Sub


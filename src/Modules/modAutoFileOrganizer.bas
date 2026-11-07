Attribute VB_Name = "modAutoFileOrganizer"
Option Explicit

Sub RunAutoFileOrganizer()

    Dim sourceFolder As String
    Dim destinationFolder As String

    Dim fileAction As VbMsgBoxResult
    Dim startTime As Double

    Dim fso As Object
    Dim mainFolder As Object
    Dim oneFile As Object

    Dim ext As String
    Dim targetFolder As String
    Dim newFilePath As String
    Dim folderCategory As String

    Dim processedCount As Long
    Dim totalFiles As Long

    Dim wbReport As Workbook
    Dim wsReport As Worksheet
    Dim nextRow As Long

    startTime = Timer

    Set fso = CreateObject("Scripting.FileSystemObject")

    '-----------------------------------
    ' STEP 1/3 - SOURCE FOLDER
    '-----------------------------------
    With Application.FileDialog(msoFileDialogFolderPicker)

        .Title = "Step 1/3: Select SOURCE folder containing files"

        If .Show <> -1 Then Exit Sub

        sourceFolder = .SelectedItems(1)

    End With

    '-----------------------------------
    ' STEP 2/3 - DESTINATION FOLDER
    '-----------------------------------
    With Application.FileDialog(msoFileDialogFolderPicker)

        .Title = "Step 2/3: Select DESTINATION folder for output"

        If .Show <> -1 Then Exit Sub

        destinationFolder = .SelectedItems(1)

    End With

    '-----------------------------------
    ' STEP 3/3 - MOVE / COPY
    '-----------------------------------
    fileAction = MsgBox( _
        "Step 3/3: How do you want to organize?" & vbCrLf & vbCrLf & _
        "YES = MOVE Files (Cut & Paste)" & vbCrLf & _
        "NO = COPY Files (Keep Originals)" & vbCrLf & _
        "CANCEL = Exit", _
        vbYesNoCancel + vbQuestion, _
        "Select Operation")

    If fileAction = vbCancel Then Exit Sub

    Set mainFolder = fso.GetFolder(sourceFolder)

    totalFiles = mainFolder.Files.count

    '-----------------------------------
    ' CREATE REPORT WORKBOOK
    '-----------------------------------
    Set wbReport = Workbooks.Add
    Set wsReport = wbReport.Sheets(1)

    wsReport.Name = "Organization Report"

    wsReport.Range("A1:E1").Value = Array( _
        "Folder Name (Category)", _
        "File Extension", _
        "File Name", _
        "Action: Open File", _
        "Action: Open Folder")

    nextRow = 2

    '-----------------------------------
    ' PROCESS FILES
    '-----------------------------------
    For Each oneFile In mainFolder.Files

        ext = LCase(fso.GetExtensionName(oneFile.Name))

        targetFolder = GetFolderByExtension(destinationFolder, ext)

        folderCategory = Mid(targetFolder, _
                             InStrRev(targetFolder, "\") + 1)

        If Not fso.FolderExists(targetFolder) Then
            fso.CreateFolder targetFolder
        End If

        newFilePath = targetFolder & "\" & oneFile.Name

        On Error Resume Next

        If fileAction = vbYes Then

            fso.MoveFile oneFile.Path, newFilePath

        Else

            fso.CopyFile oneFile.Path, newFilePath

        End If

        If Err.Number = 0 Then

            processedCount = processedCount + 1

            '---------------------------
            ' REPORT ENTRY
            '---------------------------
            wsReport.Cells(nextRow, 1).Value = folderCategory
            wsReport.Cells(nextRow, 2).Value = UCase(ext)
            wsReport.Cells(nextRow, 3).Value = oneFile.Name

            wsReport.Hyperlinks.Add _
                Anchor:=wsReport.Cells(nextRow, 4), _
                Address:=newFilePath, _
                TextToDisplay:="Click to Open File"

            wsReport.Hyperlinks.Add _
                Anchor:=wsReport.Cells(nextRow, 5), _
                Address:=targetFolder, _
                TextToDisplay:="Open Folder"

            nextRow = nextRow + 1

        End If

        Err.Clear
        On Error GoTo 0

    Next oneFile

    '-----------------------------------
    ' FORMAT REPORT
    '-----------------------------------
    With wsReport

        .Rows(1).Font.Bold = True
        .Rows(1).Interior.color = RGB(90, 90, 90)
        .Rows(1).Font.color = RGB(255, 255, 255)

        .Range("A1:E" & nextRow - 1).Borders.LineStyle = xlContinuous

        .Columns.AutoFit

        .Rows(2).Select
        ActiveWindow.FreezePanes = True

        .Range("A1:E" & nextRow - 1).AutoFilter

    End With

    '-----------------------------------
    ' SAVE REPORT
    '-----------------------------------
    Application.DisplayAlerts = False

    wbReport.SaveAs _
        destinationFolder & "\Organization Report.xlsx"

    Application.DisplayAlerts = True

    '-----------------------------------
    ' DONE MESSAGE
    '-----------------------------------
    MsgBox "Done!" & vbCrLf & vbCrLf & _
           "Files Processed: " & processedCount & _
           " / " & totalFiles & vbCrLf & _
           "Time: " & Format(Timer - startTime, "0.00") & " sec" & vbCrLf & vbCrLf & _
           "Professional Log Sheet Created.", _
           vbInformation, "Completed"

End Sub


Private Function GetFolderByExtension(basePath As String, _
                                      fileExt As String) As String

    Select Case fileExt

        Case "xlsx", "xls", "csv"
            GetFolderByExtension = basePath & "\Excel Files"

        Case "pdf"
            GetFolderByExtension = basePath & "\PDF Files"

        Case "png", "jpg", "jpeg", "gif", "bmp"
            GetFolderByExtension = basePath & "\Image Files"

        Case "doc", "docx"
            GetFolderByExtension = basePath & "\Word Files"

        Case "ppt", "pptx"
            GetFolderByExtension = basePath & "\PowerPoint Files"

        Case "zip", "rar", "7z"
            GetFolderByExtension = basePath & "\Compressed Files"

        Case "txt"
            GetFolderByExtension = basePath & "\Text Files"

        Case "sql"
            GetFolderByExtension = basePath & "\SQL Files"

        Case "py"
            GetFolderByExtension = basePath & "\Python Files"

        Case "mp4", "mkv", "avi"
            GetFolderByExtension = basePath & "\Video Files"

        Case "exe"
            GetFolderByExtension = basePath & "\Software Files"

        Case "json"
            GetFolderByExtension = basePath & "\JSON Files"

        Case "html", "htm"
            GetFolderByExtension = basePath & "\HTML Files"

        Case "js"
            GetFolderByExtension = basePath & "\JavaScript Files"

        Case Else
            GetFolderByExtension = basePath & "\Other Files"

    End Select

End Function


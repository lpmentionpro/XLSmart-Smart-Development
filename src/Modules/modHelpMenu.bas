Attribute VB_Name = "modHelpMenu"
Option Explicit

Private Const TRUSTED_ADDIN_DOWNLOAD_URL As String = _
    "https://raw.githubusercontent.com/lpmentionpro/xlsmart-smart-development/main/workbook/xlsmart_user.xlam"

#If VBA7 Then
    Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
#Else
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If

'=================================
' HOW TO USE XLSMART
'=================================
Public Sub Help_HowToUse(control As IRibbonControl)

    MsgBox "XLSmart User Guide Coming Soon.", _
           vbInformation, _
           "How to Use XLSmart"

End Sub


'=================================
' CHECK FOR UPDATES
'=================================
Public Sub Help_CheckUpdates(control As IRibbonControl)

    Dim http As Object
    Dim responseText As String
    Dim latestVersion As String
    Dim downloadUrl As String
    Dim sha256Expected As String

    On Error GoTo ErrHandler

    Set http = CreateObject("MSXML2.XMLHTTP")

    http.Open "GET", UPDATE_JSON_URL, False
    http.Send

    If http.Status <> 200 Then

        MsgBox "Unable to check updates right now." & vbCrLf & vbCrLf & _
               "Please check your internet connection.", _
               vbExclamation, _
               "XLSmart Update"

        Exit Sub

    End If

    responseText = http.responseText

    latestVersion = GetJsonValue(responseText, "version")
    downloadUrl = GetJsonValue(responseText, "download_url")
    sha256Expected = GetJsonValue(responseText, "sha256")

    If Not IsValidVersion(latestVersion) Or _
       Not IsTrustedDownloadUrl(downloadUrl) Or _
       Not IsSha256(sha256Expected) Then

        MsgBox "The update information could not be verified." & vbCrLf & vbCrLf & _
               "Your installed XLSmart add-in was not changed.", _
               vbExclamation, _
               "XLSmart Update"
        Exit Sub

    End If

    If latestVersion = APP_VERSION Then

        MsgBox "You are already using the latest version of XLSmart (" & _
               APP_VERSION & ").", _
               vbInformation, _
               "XLSmart Update"

    ElseIf Not IsVersionNewer(latestVersion, APP_VERSION) Then

        MsgBox "The update server did not provide a newer XLSmart version." & vbCrLf & vbCrLf & _
               "Your installed add-in was not changed.", _
               vbInformation, _
               "XLSmart Update"

    Else

        If MsgBox( _
            "New XLSmart update available!" & vbCrLf & vbCrLf & _
            "Current Version: " & APP_VERSION & vbCrLf & _
            "Latest Version: " & latestVersion & vbCrLf & vbCrLf & _
            "Do you want to download and install it now?" & vbCrLf & vbCrLf & _
            "Excel will close. Reopen Excel normally after the update finishes.", _
            vbYesNo + vbQuestion, _
            "XLSmart Update") = vbYes Then

            Call PerformUpdate(downloadUrl, sha256Expected)

        End If

    End If

    Exit Sub

ErrHandler:

    MsgBox "Internet connection or update server problem.", _
           vbExclamation, _
           "XLSmart Update"

End Sub


'=================================
' PERFORM AUTOMATIC UPDATE
'=================================
Private Sub PerformUpdate(ByVal downloadUrl As String, ByVal expectedSha256 As String)

    Dim http As Object
    Dim stream As Object
    Dim shell As Object
    Dim dst As String
    Dim stage As String
    Dim statusPath As String
    Dim helperPath As String
    Dim downloadedHash As String
    Dim responseBody As Variant
    Dim command As String
    Dim excelPid As Long

    On Error GoTo ErrHandler

    If HasUnsavedUserWorkbooks() Then
        MsgBox "Please save or close your unsaved workbooks before updating XLSmart." & _
               vbCrLf & vbCrLf & "No update files were installed.", _
               vbExclamation, "XLSmart Update"
        Exit Sub
    End If

    Application.StatusBar = "XLSmart: Downloading update..."

    dst = Environ$("APPDATA") & "\Microsoft\AddIns\XLSmart_User.xlam"
    excelPid = GetCurrentProcessId
    stage = dst & ".staged." & Format(Now, "yyyymmdd_hhnnss") & "_" & _
            CStr(excelPid) & "_" & Format$(CLng(Timer * 1000), "00000000") & ".xlam"
    statusPath = stage & ".log"
    helperPath = Environ$("APPDATA") & "\Microsoft\AddIns\Install-AddinUpdate.ps1"

    If Dir(helperPath) = "" Then
        Err.Raise vbObjectError + 503, , "The update helper is missing from the XLSmart installation."
    End If

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", downloadUrl, False
    http.Send

    If http.Status <> 200 Then
        Err.Raise vbObjectError + 500, , "Download failed. HTTP Status: " & http.Status
    End If
    responseBody = http.responseBody
    If IsEmpty(responseBody) Then
        Err.Raise vbObjectError + 504, , "The update server returned an empty file."
    End If

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write responseBody
    stream.SaveToFile stage, 2
    stream.Close
    Set stream = Nothing

    downloadedHash = GetFileSha256(stage)
    If LCase$(downloadedHash) <> LCase$(expectedSha256) Then
        Kill stage
        Err.Raise vbObjectError + 501, , "File integrity check failed."
    End If

    Application.StatusBar = "XLSmart: Update staged; Excel will close..."
    Set shell = CreateObject("WScript.Shell")
    command = "powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File " & _
              QuoteShellArgument(helperPath) & _
              " -StagedPath " & QuoteShellArgument(stage) & _
              " -DestinationPath " & QuoteShellArgument(dst) & _
              " -ExpectedSha256 " & QuoteShellArgument(expectedSha256) & _
              " -ExcelPid " & CStr(excelPid) & _
              " -StatusPath " & QuoteShellArgument(statusPath)

    If shell.Run(command, 0, False) <> 0 Then
        Err.Raise vbObjectError + 505, , "The external update helper could not be started."
    End If

    MsgBox "Update file downloaded and verified." & vbCrLf & vbCrLf & _
           "Excel will now close. Reopen Excel normally after a few seconds to use the new version.", _
           vbInformation, "XLSmart Update"

    Application.Quit
    Exit Sub

ErrHandler:

    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    Application.StatusBar = False
    MsgBox "Update failed:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Update Error"

End Sub


'=================================
' CALCULATE FILE SHA256
'=================================
Private Function HasUnsavedUserWorkbooks() As Boolean
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If Not wb.IsAddin Then
            If wb.Saved = False Then
                HasUnsavedUserWorkbooks = True
                Exit Function
            End If
        End If
    Next wb
End Function

Private Function GetFileSha256(ByVal filePath As String) As String

    Dim shell As Object
    Dim exec As Object
    Dim output As String
    Dim command As String
    Dim line As Variant

    On Error GoTo HashError

    Set shell = CreateObject("WScript.Shell")
    command = "cmd.exe /d /c certutil -hashfile " & QuoteShellArgument(filePath) & " SHA256"
    Set exec = shell.Exec(command)
    output = exec.StdOut.ReadAll()

    For Each line In Split(Replace$(output, vbCr, ""), vbLf)
        line = Replace$(Trim$(CStr(line)), " ", "")
        If Len(line) = 64 And IsSha256(CStr(line)) Then
            GetFileSha256 = UCase$(CStr(line))
            Exit Function
        End If
    Next line
    Err.Raise vbObjectError + 506, , "Unable to calculate the update file SHA-256."

HashError:
    GetFileSha256 = vbNullString

End Function

Private Function QuoteShellArgument(ByVal value As String) As String
    ' Windows file names cannot contain quotation marks. Quoting makes spaces safe.
    QuoteShellArgument = """" & value & """"
End Function

Private Function IsSha256(ByVal value As String) As Boolean
    Dim i As Long
    Dim ch As String

    If Len(value) <> 64 Then Exit Function
    For i = 1 To 64
        ch = Mid$(value, i, 1)
        If InStr(1, "0123456789abcdefABCDEF", ch, vbBinaryCompare) = 0 Then Exit Function
    Next i
    IsSha256 = True
End Function

Private Function IsValidVersion(ByVal value As String) As Boolean
    Dim parts() As String
    Dim i As Long

    parts = Split(value, ".")
    If UBound(parts) <> 2 Then Exit Function
    For i = 0 To 2
        If Len(parts(i)) = 0 Or Not IsNumeric(parts(i)) Then Exit Function
        If CStr(CLng(parts(i))) <> parts(i) Then Exit Function
    Next i
    IsValidVersion = True
End Function

Private Function IsVersionNewer(ByVal candidate As String, ByVal current As String) As Boolean
    Dim newParts() As String
    Dim oldParts() As String
    Dim i As Long

    newParts = Split(candidate, ".")
    oldParts = Split(current, ".")
    For i = 0 To 2
        If CLng(newParts(i)) > CLng(oldParts(i)) Then
            IsVersionNewer = True
            Exit Function
        ElseIf CLng(newParts(i)) < CLng(oldParts(i)) Then
            Exit Function
        End If
    Next i
End Function

Private Function IsTrustedDownloadUrl(ByVal value As String) As Boolean
    IsTrustedDownloadUrl = _
        (StrComp(value, TRUSTED_ADDIN_DOWNLOAD_URL, vbTextCompare) = 0)
End Function


Private Function GetJsonValue(jsonText As String, key As String) As String

    Dim startPos As Long
    Dim endPos As Long
    Dim searchKey As String
    Dim cleanText As String

    ' remove line breaks and spaces
    cleanText = Replace(jsonText, vbCr, "")
    cleanText = Replace(cleanText, vbLf, "")
    cleanText = Replace(cleanText, " ", "")

    searchKey = """" & key & """:"""

    startPos = InStr(cleanText, searchKey)

    If startPos = 0 Then Exit Function

    startPos = startPos + Len(searchKey)

    If Mid(cleanText, startPos, 1) = """" Then
        startPos = startPos + 1
    End If

    endPos = InStr(startPos, cleanText, """")

    If endPos > startPos Then
        GetJsonValue = Mid(cleanText, startPos, endPos - startPos)
    End If

End Function


'=================================
' LICENSE STATUS
'=================================
Public Sub Help_LicenseStatus(control As IRibbonControl)

    MsgBox _
        "XLSmart License Information" & vbCrLf & vbCrLf & _
        "Version: " & APP_VERSION & vbCrLf & _
        "Edition: " & APP_EDITION & vbCrLf & _
        "Status: Active", _
        vbInformation, _
        "License Status"

End Sub


'=================================
' UPGRADE TO PRO
'=================================
Public Sub Help_UpgradePro(control As IRibbonControl)

    MsgBox "XLSmart Pro Coming Soon.", _
           vbInformation, _
           "Upgrade XLSmart"

End Sub


'=================================
' UNINSTALL XLSMART
'=================================
Public Sub Help_Uninstall(control As IRibbonControl)

    Dim ans As VbMsgBoxResult
    Dim i As Long
    Dim addinPath As String
    Dim wbName As String

    On Error GoTo ErrHandler

    wbName = "XLSmart_User"

    ans = MsgBox( _
        "Are you sure you want to uninstall XLSmart?" & vbCrLf & vbCrLf & _
        "Excel will restart after uninstall.", _
        vbYesNo + vbQuestion, _
        "Uninstall XLSmart")

    If ans = vbNo Then Exit Sub

    addinPath = Environ$("APPDATA") & _
                "\Microsoft\AddIns\XLSmart_User.xlam"

    Application.DisplayAlerts = False

    ' disable addin
    For i = AddIns.count To 1 Step -1

        If LCase(AddIns(i).Name) = _
           LCase("XLSmart_User.xlam") Then

            AddIns(i).Installed = False

            ' remove from addins list
            AddIns(i).Delete

            Exit For

        End If

    Next i

    ' delete physical file
    If Dir(addinPath) <> "" Then
        Kill addinPath
    End If

    MsgBox "XLSmart removed successfully." & vbCrLf & _
           "Excel will now close.", _
           vbInformation, _
           "XLSmart"

    Application.Quit

    Exit Sub

ErrHandler:

    MsgBox "Uninstall failed:" & vbCrLf & _
           Err.Description, _
           vbCritical, _
           "XLSmart"

End Sub


'=================================
' WEBSITE
'=================================
Public Sub Help_Website(control As IRibbonControl)

    ThisWorkbook.FollowHyperlink _
    "https://yourwebsite.com"

End Sub


'=================================
' YOUTUBE
'=================================
Public Sub Help_YouTube(control As IRibbonControl)

    ThisWorkbook.FollowHyperlink _
    "https://youtube.com"

End Sub


'=================================
' LINKEDIN
'=================================
Public Sub Help_LinkedIn(control As IRibbonControl)

    ThisWorkbook.FollowHyperlink _
    "https://linkedin.com"

End Sub


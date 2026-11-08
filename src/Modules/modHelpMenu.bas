Attribute VB_Name = "modHelpMenu"
Option Explicit

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

    If latestVersion = APP_VERSION Then

        MsgBox "You are already using the latest version of XLSmart (" & _
               APP_VERSION & ").", _
               vbInformation, _
               "XLSmart Update"

    Else

        If MsgBox( _
            "New XLSmart update available!" & vbCrLf & vbCrLf & _
            "Current Version: " & APP_VERSION & vbCrLf & _
            "Latest Version: " & latestVersion & vbCrLf & vbCrLf & _
            "Do you want to update now?" & vbCrLf & vbCrLf & _
            "(Excel will restart after update)", _
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
    Dim dst As String
    Dim ai As AddIn
    Dim downloadedHash As String

    On Error GoTo ErrHandler

    Application.DisplayAlerts = False
    Application.StatusBar = "XLSmart: Downloading update..."

    ' Download URL
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", downloadUrl, False
    http.Send

    If http.Status <> 200 Then
        Err.Raise vbObjectError + 500, , "Download failed. HTTP Status: " & http.Status
    End If

    Application.StatusBar = "XLSmart: Installing update..."

    ' Destination path
    dst = Environ$("APPDATA") & "\Microsoft\AddIns\XLSmart_User.xlam"

    ' Backup existing version
    If Dir(dst) <> "" Then
        On Error Resume Next
        SetAttr dst, vbNormal
        Dim backupPath As String
        backupPath = dst & ".backup"
        Kill backupPath
        Name dst As backupPath
        On Error GoTo ErrHandler
    End If

    ' Save downloaded file
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile dst, 2
    stream.Close

    ' Verify SHA256 if provided
    If Len(expectedSha256) > 0 Then
        downloadedHash = GetFileSha256(dst)
        If LCase(downloadedHash) <> LCase(expectedSha256) Then
            Err.Raise vbObjectError + 501, , "File integrity check failed."
        End If
    End If

    ' Unload old add-in
    On Error Resume Next
    For Each ai In AddIns
        If LCase(ai.Name) = "xlsmart_user.xlam" Then
            ai.Installed = False
            AddIns.Remove ai
            Exit For
        End If
    Next ai
    On Error GoTo ErrHandler

    ' Install new add-in
    Set ai = AddIns.Add(dst)
    ai.Installed = True

    Application.StatusBar = False
    Application.DisplayAlerts = True

    MsgBox "XLSmart updated successfully!" & vbCrLf & vbCrLf & _
           "Excel will now restart.", _
           vbInformation, _
           "Update Complete"

    Application.Quit

    Exit Sub

ErrHandler:

    Application.StatusBar = False
    Application.DisplayAlerts = True

    MsgBox "Update failed:" & vbCrLf & vbCrLf & _
           Err.Description, _
           vbCritical, _
           "Update Error"

End Sub


'=================================
' CALCULATE FILE SHA256
'=================================
Private Function GetFileSha256(ByVal filePath As String) As String

    Dim shell As Object
    Dim exec As Object
    Dim output As String

    On Error Resume Next

    Set shell = CreateObject("WScript.Shell")
    Set exec = shell.Exec("powershell -NoProfile -Command ""(Get-FileHash -Path '" & filePath & "' -Algorithm SHA256).Hash""")
    output = exec.StdOut.ReadAll()

    GetFileSha256 = Trim(output)

    On Error GoTo 0

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


Attribute VB_Name = "Module1"

Option Explicit

Sub ExportAllVBA()

    Dim vbComp As Object
    Dim folderPath As String
    Dim subFolder As String
    Dim fileExt As String
    Dim fso As Object

    folderPath = ThisWorkbook.Path & "\VBA_Export"

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If

    For Each vbComp In ThisWorkbook.VBProject.VBComponents

        Select Case vbComp.Type

            Case 1
                subFolder = "Modules"
                fileExt = ".bas"

            Case 2
                subFolder = "Classes"
                fileExt = ".cls"

            Case 3
                subFolder = "Forms"
                fileExt = ".frm"

            Case 100
                subFolder = "Documents"
                fileExt = ".cls"

            Case Else
                GoTo NextComponent

        End Select

        If Not fso.FolderExists(folderPath & "\" & subFolder) Then
            fso.CreateFolder folderPath & "\" & subFolder
        End If

        If vbComp.Type = 100 Then
            Debug.Print "Document Module: " & vbComp.Name
        Else
            vbComp.Export folderPath & "\" & subFolder & "\" & _
                          vbComp.Name & fileExt
        End If

        Debug.Print "Exported: " & vbComp.Name

NextComponent:
    Next vbComp

    MsgBox "VBA Export Completed!" & vbCrLf & folderPath, vbInformation

End Sub



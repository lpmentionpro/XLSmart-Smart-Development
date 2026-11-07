Attribute VB_Name = "modExportAll"
Option Explicit

Public Sub ExportAllCode()

    Dim vbComp As Object
    Dim exportPath As String
    Dim ext As String

    exportPath = ThisWorkbook.Path & "\XLSmart_Source\"

    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If

    For Each vbComp In ThisWorkbook.VBProject.VBComponents

        Select Case vbComp.Type

            Case 1
                ext = ".bas"

            Case 2
                ext = ".cls"

            Case 3
                ext = ".frm"

            Case 100
                ext = ".cls"

            Case Else
                ext = ""

        End Select

        If ext <> "" Then

            On Error Resume Next
            Kill exportPath & vbComp.Name & ext
            On Error GoTo 0

            vbComp.Export exportPath & vbComp.Name & ext

        End If

    Next vbComp

    MsgBox "Done!" & vbCrLf & exportPath

End Sub

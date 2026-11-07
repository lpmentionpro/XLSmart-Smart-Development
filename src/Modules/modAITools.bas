Attribute VB_Name = "modAITools"
Option Explicit

' ============================================================
' AI FORMULA GENERATOR
' ============================================================
Public Sub AIFormulaGenerator(control As IRibbonControl)

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Collect sheet column headers
    Dim sheetInfo As String
    sheetInfo = GetSheetHeaders(ws)

    Dim userPrompt As String
    userPrompt = InputBox( _
        "Describe what formula you need:" & vbNewLine & vbNewLine & _
        "Examples:" & vbNewLine & _
        "  - Sum of total sales" & vbNewLine & _
        "  - Convert DOB to exact age in years months days" & vbNewLine & _
        "  - VLOOKUP customer name from Sheet2" & vbNewLine & _
        "  - Count orders where state is Maharashtra" & vbNewLine & vbNewLine & _
        "Sheet: " & ws.Name & vbNewLine & _
        "Columns: " & sheetInfo, _
        "XLSmart - AI Formula Generator")

    If Trim(userPrompt) = "" Then Exit Sub

    Dim prompt As String
    prompt = "I have an Excel sheet named " & ws.Name & _
             " with columns: " & sheetInfo & ". " & _
             "Task: " & userPrompt & ". " & _
             "Give me only the Excel formula starting with = sign. No explanation."

    OpenChatGPT prompt
End Sub

' ============================================================
' AI FORMULA EXPLAINER
' ============================================================
Public Sub AIFormulaExplainer(control As IRibbonControl)

    If Selection.Cells.count > 1 Then
        MsgBox "Please select a single cell containing a formula.", vbExclamation, "XLSmart"
        Exit Sub
    End If

    Dim cell As Range
    Set cell = Selection

    If left(cell.Formula, 1) <> "=" Then
        MsgBox "Selected cell does not contain a formula!" & vbNewLine & _
               "Please select a cell with a formula (starts with =).", _
               vbExclamation, "XLSmart"
        Exit Sub
    End If

    Dim lang As String
    lang = InputBox("Explain formula in which language?", "XLSmart", "Hindi")
    If lang = "" Then lang = "Hindi"

    Dim prompt As String
    prompt = "Explain this Excel formula step by step in " & lang & ": " & cell.Formula & _
             ". Break down each part clearly."

    OpenChatGPT prompt
End Sub

' ============================================================
' AI FORMULA ERROR DETECTOR
' ============================================================
Public Sub AIFormulaErrorDetector(control As IRibbonControl)

    If Selection.Cells.count > 1 Then
        MsgBox "Please select a single cell containing a formula error.", vbExclamation, "XLSmart"
        Exit Sub
    End If

    Dim cell As Range
    Set cell = Selection

    If left(cell.Formula, 1) <> "=" Then
        MsgBox "Selected cell does not contain a formula!", vbExclamation, "XLSmart"
        Exit Sub
    End If

    Dim errVal As String
    errVal = CStr(cell.Value)

    Dim prompt As String
    prompt = "This Excel formula has an error. Find the exact reason and give corrected formula." & _
             " Formula: " & cell.Formula & _
             " Error shown: " & errVal & _
             ". Reply with: 1) Error reason 2) Corrected formula only."

    OpenChatGPT prompt
End Sub

' ============================================================
' AI COMMENT GENERATOR
' ============================================================
Public Sub AICommentGenerator(control As IRibbonControl)

    If Selection.Cells.count > 1 Then
        MsgBox "Please select a single cell to generate comment for.", vbExclamation, "XLSmart"
        Exit Sub
    End If

    Dim cell As Range
    Set cell = Selection

    Dim cellValue As String
    cellValue = CStr(cell.Value)

    If cellValue = "" Then
        MsgBox "Selected cell is empty!", vbExclamation, "XLSmart"
        Exit Sub
    End If

    Dim commentType As String
    commentType = InputBox( _
        "What kind of comment for cell " & cell.Address & "?" & vbNewLine & vbNewLine & _
        "Value: " & cellValue & vbNewLine & vbNewLine & _
        "Examples:" & vbNewLine & _
        "  - Explain this value" & vbNewLine & _
        "  - Add a warning note" & vbNewLine & _
        "  - Explain this formula result", _
        "XLSmart - AI Comment Generator", _
        "Explain this value in 1-2 lines")

    If Trim(commentType) = "" Then Exit Sub

    Dim prompt As String
    prompt = "Excel cell " & cell.Address & " in sheet " & cell.Parent.Name & _
             " has value: " & cellValue & ". Task: " & commentType & _
             ". Write a short cell comment in max 2 lines. Be concise."

    OpenChatGPT prompt

    ' Let user paste comment manually
    Dim addNow As Integer
    addNow = MsgBox("ChatGPT opened in browser." & vbNewLine & vbNewLine & _
                    "Copy the AI response, then click Yes to paste as cell comment.", _
                    vbYesNo + vbQuestion, "XLSmart")

    If addNow = vbYes Then
        Dim commentText As String
        commentText = InputBox("Paste the comment for cell " & cell.Address & ":", "XLSmart")
        If Trim(commentText) <> "" Then
            On Error Resume Next
            cell.Comment.Delete
            On Error GoTo 0
            cell.AddComment commentText
            cell.Comment.Shape.TextFrame.AutoSize = True
            MsgBox "Comment added to " & cell.Address, vbInformation, "XLSmart"
        End If
    End If
End Sub

' ============================================================
' HELPER: Open ChatGPT in browser with pre-filled prompt
' ============================================================
Private Sub OpenChatGPT(prompt As String)
    Dim encoded As String
    encoded = URLEncode(prompt)
    Dim url As String
    url = "https://chat.openai.com/?q=" & encoded
    Shell "cmd /c start " & Chr(34) & Chr(34) & " " & Chr(34) & url & Chr(34), vbHide
End Sub

' ============================================================
' HELPER: URL Encode
' ============================================================
Private Function URLEncode(txt As String) As String
    Dim result As String
    result = ""
    Dim i As Integer
    Dim c As String
    For i = 1 To Len(txt)
        c = Mid(txt, i, 1)
        Select Case asc(c)
            Case 65 To 90, 97 To 122, 48 To 57
                result = result & c
            Case 32
                result = result & "+"
            Case Else
                result = result & "%" & right("0" & Hex(asc(c)), 2)
        End Select
    Next i
    If Len(result) > 2000 Then result = left(result, 2000)
    URLEncode = result
End Function

' ============================================================
' HELPER: Get sheet column headers as string
' ============================================================
Public Function GetSheetHeaders(ws As Worksheet) As String
    Dim result As String
    result = ""
    Dim j As Integer
    On Error Resume Next
    For j = 1 To ws.UsedRange.Columns.count
        Dim h As String
        h = CStr(ws.Cells(1, j).Value)
        If h <> "" Then
            result = result & h & " (Col " & j & "), "
        End If
    Next j
    On Error GoTo 0
    If Len(result) > 2 Then result = left(result, Len(result) - 2)
    GetSheetHeaders = result
End Function

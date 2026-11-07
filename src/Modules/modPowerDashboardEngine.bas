Attribute VB_Name = "modPowerDashboardEngine"
Option Explicit



Public Sub CreateDashboard(ByVal dashTitle As String, _
                           ByVal theme As String, _
                           ByVal currencyCode As String, _
                           ByVal isMulti As Boolean, _
                           ByVal rangeOrSheets As String, _
                           ByVal startRow As Long, _
                           ByVal smartLock As Boolean, _
                           ByVal gridlinesShow As Boolean)
    On Error GoTo CleanFail
    Application.StatusBar = "Power Dashboard: reading data..."

    Dim dataWb As Workbook
    Dim jsonData As String
    Dim sheetsJson As String
    Dim rng As Range

    If isMulti Then
        Set dataWb = ActiveWorkbook
        If dataWb Is Nothing Then Err.Raise vbObjectError + 100, , "No active workbook found."
        If dataWb Is ThisWorkbook Then
            Err.Raise vbObjectError + 101, , "Please activate the workbook that contains your data, then run Power Dashboard again."
        End If
        jsonData = MultiSheetJSON(dataWb, rangeOrSheets, startRow)
        sheetsJson = MultiSheetsJSON(dataWb, rangeOrSheets, startRow)
    Else
        Set rng = ResolveRange(rangeOrSheets)
        If rng Is Nothing Then Err.Raise vbObjectError + 102, , "Invalid data range: " & rangeOrSheets
        jsonData = SingleRangeJSON(rng)
        sheetsJson = SingleSheetsJSON(rng)
    End If

    If Len(sheetsJson) = 0 Then sheetsJson = "[]"

    If Len(jsonData) = 0 Or jsonData = "{}" Or jsonData = "[]" Then
        Err.Raise vbObjectError + 103, , "No readable data found."
    End If

    Application.StatusBar = "Power Dashboard: creating HTML..."
    Dim html As String
    html = BuildDashboardHTML(jsonData, sheetsJson, dashTitle, theme, currencyCode, gridlinesShow, smartLock)

    Dim htmlPath As String
    htmlPath = Environ$("TEMP") & "\power_dashboard.html"
    WriteUtf8 htmlPath, html
    If Dir(htmlPath) = "" Then Err.Raise vbObjectError + 104, , "Dashboard HTML file could not be created."

    Application.StatusBar = "Opening dashboard..."
    On Error Resume Next
    ThisWorkbook.FollowHyperlink htmlPath
    If Err.Number <> 0 Then
        Err.Clear
        Shell "cmd /c start """" """ & htmlPath & """", vbHide
    End If
    On Error GoTo CleanFail

    Application.StatusBar = False
    MsgBox "Dashboard created successfully!", vbInformation, "Power Dashboard"
    Exit Sub

CleanFail:
    Application.StatusBar = False
    MsgBox Err.Description, vbCritical, "Power Dashboard"
End Sub

' ---------------------------------------------------------------
'  HTML build: template comes from modDashTemplate (generated)
' ---------------------------------------------------------------
Private Function BuildDashboardHTML(ByVal jsonData As String, _
                                    ByVal sheetsJson As String, _
                                    ByVal dashTitle As String, _
                                    ByVal theme As String, _
                                    ByVal currencyCode As String, _
                                    ByVal gridlinesShow As Boolean, _
                                    ByVal smartLock As Boolean) As String
    Dim s As String
    Dim cfg As String
    Dim titleJs As String
    Dim p As Long

    s = GetDashboardTemplate()

    ' Config injected into the template (/*--AICONFIG_START--*/ .. END--*/).
    ' currency  -> currencySym in JS  ('' | taka | rupee | euro | dollar)
    ' showGrid  -> gridlines Show/Hide from the form
    ' smartLock -> asks for confirmation before a KPI/chart/slicer is deleted
    ' sheets    -> Data View tabs (top 50 rows of every selected sheet)
    cfg = "let aiConfig={title:" & Replace(JsonQuote(dashTitle), "</", "<\/") & ",kpis:[],charts:[],slicers:[],tables:[]," & _
          "sheets:" & Replace(sheetsJson, "</", "<\/") & "," & _
          "logoImage:'',logoIcon:''," & _
          "github:{owner:'',repo:'',branch:'main',token:'',baseUrl:''}," & _
          "theme:" & JsonQuote(TemplateThemeName(theme)) & "," & _
          "currency:" & CurrencyJsLiteral(currencyCode) & "," & _
          "showGrid:" & IIf(gridlinesShow, "true", "false") & "," & _
          "smartLock:" & IIf(smartLock, "true", "false") & "};"

    s = ReplaceBetween(s, "/*--RAWDATA_START--*/", "/*--RAWDATA_END--*/", _
                       "let rawData=" & Replace(jsonData, "</", "<\/") & ";")
    s = ReplaceBetween(s, "/*--AICONFIG_START--*/", "/*--AICONFIG_END--*/", cfg)

    titleJs = "<script>(function(){var t=" & Replace(JsonQuote(dashTitle), "</", "<\/") & _
              ";document.title=t;var e=document.getElementById('mainTitle');if(e)e.innerText=t;})();</script>"
    p = InStrRev(s, "</body>", -1, vbTextCompare)
    If p > 0 Then
        s = left$(s, p - 1) & titleJs & Mid$(s, p)
    Else
        s = s & titleJs
    End If

    BuildDashboardHTML = s
End Function

' Currency as a JS string literal built from \uXXXX escapes, so this .bas file
' stays pure ASCII (VBE import would corrupt real Taka/Rupee characters).
Private Function CurrencyJsLiteral(ByVal code As String) As String
    Select Case UCase$(Trim$(code))
        Case "BDT":    CurrencyJsLiteral = """" & "\u09F3" & """"   ' BDT taka
        Case "INR":    CurrencyJsLiteral = """" & "\u20B9" & """"   ' INR rupee
        Case "EUR":    CurrencyJsLiteral = """" & "\u20AC" & """"   ' EUR euro
        Case "USD":    CurrencyJsLiteral = """" & "$" & """"        ' USD dollar
        Case Else:     CurrencyJsLiteral = """" & "" & """"         ' None
    End Select
End Function

Private Function ReplaceBetween(ByVal src As String, ByVal startMarker As String, _
                                ByVal endMarker As String, ByVal replacement As String) As String
    Dim a As Long, b As Long
    a = InStr(1, src, startMarker, vbBinaryCompare)
    If a = 0 Then Err.Raise vbObjectError + 107, , "Template marker missing: " & startMarker
    b = InStr(a + Len(startMarker), src, endMarker, vbBinaryCompare)
    If b = 0 Then Err.Raise vbObjectError + 108, , "Template marker missing: " & endMarker
    ReplaceBetween = left$(src, a + Len(startMarker) - 1) & vbLf & replacement & vbLf & Mid$(src, b)
End Function

' Form theme -> template theme name (must match liveThemes keys in Sec19_Themes)
Private Function TemplateThemeName(ByVal theme As String) As String
    Select Case LCase$(Trim$(theme))
        Case "ocean":       TemplateThemeName = "Ocean"
        Case "midnight":    TemplateThemeName = "Midnight"
        Case "slate":       TemplateThemeName = "Dark"       ' Slate (Dark Blue)
        Case "forest":      TemplateThemeName = "Forest"     ' Forest (Green)
        Case "rose":        TemplateThemeName = "Rose"       ' Rose (Pink)
        Case "navy":        TemplateThemeName = "Navy"       ' Navy (Deep Blue)
        Case "dark":        TemplateThemeName = "Dark"
        Case "light":       TemplateThemeName = "Light"
        Case "animated":    TemplateThemeName = "Animated"
        Case Else:          TemplateThemeName = "Midnight"   ' matches the form default
    End Select
End Function

' UTF-8 writer (Print # would corrupt Bengali text and currency symbols)
Private Sub WriteUtf8(ByVal filePath As String, ByVal text As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText text
    st.SaveToFile filePath, 2
    st.Close
End Sub

' ---------------------------------------------------------------
'  Data reading
' ---------------------------------------------------------------
Private Function ResolveRange(ByVal addressText As String) As Range
    On Error GoTo failSafe
    Dim parts() As String
    Dim wb As Workbook
    parts = Split(addressText, "|")
    If UBound(parts) <> 2 Then Exit Function
    On Error Resume Next
    Set wb = Workbooks(parts(0))
    On Error GoTo failSafe
    If wb Is Nothing Then Exit Function
    Set ResolveRange = wb.Worksheets(parts(1)).Range(parts(2))
    Exit Function
failSafe:
    Set ResolveRange = Nothing
End Function

Private Function SingleRangeJSON(ByVal rng As Range) As String
    If rng Is Nothing Then Exit Function
    If rng.Rows.count < 2 Or rng.Columns.count < 1 Then Exit Function
    SingleRangeJSON = RangeMatrixJSON(rng)
End Function

Private Function MultiSheetJSON(ByVal wb As Workbook, ByVal sheetList As String, ByVal startRow As Long) As String
    Dim parts() As String
    Dim i As Long, idx As Long
    Dim ws As Worksheet
    Dim lastCol As Long, lastRow As Long

    If startRow < 1 Then startRow = 1
    parts = Split(sheetList, ",")
    For i = LBound(parts) To UBound(parts)
        idx = CLng(val(Trim$(parts(i))))
        If idx >= 1 And idx <= wb.Worksheets.count Then
            Set ws = wb.Worksheets(idx)
            lastCol = ws.Cells(startRow, ws.Columns.count).End(xlToLeft).Column
            lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
            If lastCol >= 1 And lastRow > startRow Then
                ' Template takes one raw-data matrix: first non-empty selected sheet
                MultiSheetJSON = RangeMatrixJSON(ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, lastCol)))
                Exit Function
            End If
        End If
    Next i
    MultiSheetJSON = "[]"
End Function

' Every selected sheet -> [{"name":"Sales","rows":[[header],[row],...]}, ...]
' Data View modal ei array theke sheet tab gulo banay (header + top 50 row).
Private Function MultiSheetsJSON(ByVal wb As Workbook, ByVal sheetList As String, ByVal startRow As Long) As String
    Dim parts() As String
    Dim i As Long, idx As Long
    Dim ws As Worksheet
    Dim lastCol As Long, lastRow As Long
    Dim entries() As String
    Dim n As Long

    If startRow < 1 Then startRow = 1
    parts = Split(sheetList, ",")
    For i = LBound(parts) To UBound(parts)
        idx = CLng(val(Trim$(parts(i))))
        If idx >= 1 And idx <= wb.Worksheets.count Then
            Set ws = wb.Worksheets(idx)
            lastCol = ws.Cells(startRow, ws.Columns.count).End(xlToLeft).Column
            lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
            If lastCol >= 1 And lastRow > startRow Then
                n = n + 1
                ReDim Preserve entries(1 To n)
                entries(n) = SheetEntryJSON(ws.Name, _
                    RangeMatrixJSON(ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, lastCol)), 51))
            End If
        End If
    Next i
    If n = 0 Then
        MultiSheetsJSON = "[]"
    Else
        MultiSheetsJSON = "[" & Join(entries, ",") & "]"
    End If
End Function

' Single range (workbook-er ekta sheet) -> ekta-i Data View tab
Private Function SingleSheetsJSON(ByVal rng As Range) As String
    If rng Is Nothing Then
        SingleSheetsJSON = "[]"
        Exit Function
    End If
    SingleSheetsJSON = "[" & SheetEntryJSON(rng.Parent.Name, RangeMatrixJSON(rng, 51)) & "]"
End Function

Private Function SheetEntryJSON(ByVal sheetName As String, ByVal rowsJson As String) As String
    SheetEntryJSON = "{""name"":" & JsonQuote(sheetName) & ",""rows"":" & rowsJson & "}"
End Function

Private Function RangeMatrixJSON(ByVal rng As Range, Optional ByVal maxRows As Long = 0) As String
    Dim v As Variant
    Dim nR As Long, nC As Long, r As Long, c As Long
    Dim rowsArr() As String
    Dim cellsArr() As String

    v = rng.Value
    nR = UBound(v, 1)
    nC = UBound(v, 2)
    If maxRows > 0 And nR > maxRows Then nR = maxRows
    ReDim rowsArr(1 To nR)
    ReDim cellsArr(1 To nC)

    For r = 1 To nR
        For c = 1 To nC
            If r = 1 Then
                cellsArr(c) = JsonQuote(HeaderText(v(1, c), c))
            Else
                cellsArr(c) = JsonValue(v(r, c))
            End If
        Next c
        rowsArr(r) = "[" & Join(cellsArr, ",") & "]"
    Next r

    RangeMatrixJSON = "[" & Join(rowsArr, ",") & "]"
End Function

Private Function HeaderText(ByVal v As Variant, ByVal c As Long) As String
    Dim t As String
    If Not IsError(v) Then t = Trim$(CStr(v))
    If Len(t) = 0 Then t = "Column" & c
    HeaderText = t
End Function

Private Function JsonValue(ByVal v As Variant) As String
    Select Case VarType(v)
        Case vbEmpty, vbNull, vbError
            JsonValue = "null"
        Case vbDate
            If CDbl(v) = Int(CDbl(v)) Then
                JsonValue = JsonQuote(Format$(CDate(v), "yyyy-mm-dd"))
            Else
                JsonValue = JsonQuote(Format$(CDate(v), "yyyy-mm-dd hh:nn:ss"))
            End If
        Case vbBoolean
            JsonValue = IIf(v, "true", "false")
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            JsonValue = Trim$(Str$(CDbl(v)))
        Case Else
            JsonValue = JsonQuote(CStr(v))
    End Select
End Function

Private Function JsonQuote(ByVal text As String) As String
    text = Replace(text, Chr$(92), Chr$(92) & Chr$(92))
    text = Replace(text, Chr$(34), Chr$(92) & Chr$(34))
    text = Replace(text, vbCrLf, "\n")
    text = Replace(text, vbCr, "\n")
    text = Replace(text, vbLf, "\n")
    text = Replace(text, vbTab, "\t")
    JsonQuote = """" & text & """"
End Function

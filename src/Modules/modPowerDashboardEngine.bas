Attribute VB_Name = "modPowerDashboardEngine"
Option Explicit

Private Const TEMPLATE_FILE_NAME As String = "PowerDashboardTemplate.html"
Private Const TEMPLATE_ASSET_FOLDER As String = "assets"
Private Const TEMPLATE_URL As String = "https://raw.githubusercontent.com/lpmentionpro/XLSmart-Smart-Development/main/assets/PowerDashboardTemplate.html"
Private Const TEMPLATE_START_RAW As String = "/*--RAWDATA_START--*/"
Private Const TEMPLATE_END_RAW As String = "/*--RAWDATA_END--*/"
Private Const TEMPLATE_START_CONFIG As String = "/*--AICONFIG_START--*/"
Private Const TEMPLATE_END_CONFIG As String = "/*--AICONFIG_END--*/"
Private Const DASHBOARD_SCHEMA_VERSION As Long = 2
Private Const DASHBOARD_APPLICATION_NAME As String = "XLSmart"
Private Const DASHBOARD_DEFAULT_THEME As String = "Animated"
Private Const DASHBOARD_BOOTSTRAP_FILE_NAME As String = "XLSmart_PowerDashboard.html"

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
    Dim rawDataJson As String

    If isMulti Then
        Set dataWb = ActiveWorkbook
        If dataWb Is Nothing Then
            Err.Raise vbObjectError + 100, , "No active workbook found."
        End If
        If dataWb Is ThisWorkbook Then
            Err.Raise vbObjectError + 101, , "Please activate the workbook that contains your data, then run Power Dashboard again."
        End If
        rawDataJson = MultiSheetRawDataJSON(dataWb, rangeOrSheets, startRow)
    Else
        Dim rng As Range
        Set rng = ResolveRange(rangeOrSheets)
        If rng Is Nothing Then
            Err.Raise vbObjectError + 102, , "Invalid data range: " & rangeOrSheets
        End If
        Set dataWb = rng.Worksheet.Parent
        rawDataJson = RangeRawDataJSON(rng)
    End If

    If Len(rawDataJson) = 0 Or rawDataJson = "[]" Then
        Err.Raise vbObjectError + 103, , "No readable data found. Select a header row and at least one data row."
    End If

    Application.StatusBar = "Power Dashboard: loading professional template..."

    Dim templateHtml As String
    templateHtml = LoadDashboardTemplate()
    If Len(templateHtml) = 0 Then
        Err.Raise vbObjectError + 105, , "Power Dashboard template is missing. Reinstall XLSmart so " & TEMPLATE_FILE_NAME & " is copied to the AddIns folder."
    End If

    Dim configJson As String
    configJson = BuildDashboardConfigJSON(dashTitle, theme, currencyCode, smartLock, gridlinesShow, rawDataJson)

    Dim htmlContent As String
    htmlContent = ReplaceTemplateSection(templateHtml, TEMPLATE_START_RAW, TEMPLATE_END_RAW, _
                    "let rawData = " & rawDataJson & ";")
    htmlContent = ReplaceTemplateSection(htmlContent, TEMPLATE_START_CONFIG, TEMPLATE_END_CONFIG, _
                    "let aiConfig = " & configJson & ";")
    htmlContent = ApplyDashboardSettings(htmlContent, dashTitle, theme, currencyCode, smartLock, gridlinesShow)

    Dim htmlPath As String
    htmlPath = Environ$("TEMP") & "\" & DASHBOARD_BOOTSTRAP_FILE_NAME
    WriteUtf8 htmlPath, htmlContent

    If Dir$(htmlPath) = "" Then
        Err.Raise vbObjectError + 104, , "Dashboard HTML file could not be created."
    End If

    Application.StatusBar = "Opening dashboard..."
    OpenDashboard htmlPath

    Application.StatusBar = False
    MsgBox "Professional Power Dashboard created successfully!", vbInformation, "Power Dashboard"
    Exit Sub

CleanFail:
    Application.StatusBar = False
    MsgBox Err.Description, vbCritical, "Power Dashboard"
End Sub

Private Function ResolveRange(ByVal addressText As String) As Range
    On Error GoTo failSafe
    Dim parts() As String
    Dim wb As Workbook
    Dim ws As Worksheet

    parts = Split(addressText, "|")
    If UBound(parts) <> 2 Then Exit Function

    Set wb = Workbooks(parts(0))
    Set ws = wb.Worksheets(parts(1))
    Set ResolveRange = ws.Range(parts(2))
    Exit Function

failSafe:
    Set ResolveRange = Nothing
End Function

Private Function RangeRawDataJSON(ByVal rng As Range) As String
    If rng Is Nothing Then Exit Function
    If rng.Rows.Count < 2 Or rng.Columns.Count < 1 Then Exit Function

    Dim result As String
    Dim r As Long
    result = "["

    For r = 1 To rng.Rows.Count
        If r > 1 Then result = result & ","
        result = result & RangeRowJSON(rng.Rows(r))
    Next r

    RangeRawDataJSON = result & "]"
End Function

Private Function MultiSheetRawDataJSON(ByVal wb As Workbook, ByVal sheetList As String, ByVal startRow As Long) As String
    Dim parts() As String
    Dim i As Long
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim lastRow As Long
    Dim dataRng As Range

    parts = Split(sheetList, ",")
    For i = LBound(parts) To UBound(parts)
        If Val(Trim$(parts(i))) >= 1 And Val(Trim$(parts(i))) <= wb.Worksheets.Count Then
            Set ws = wb.Worksheets(Val(Trim$(parts(i))))
            lastCol = ws.Cells(startRow, ws.Columns.Count).End(xlToLeft).Column
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

            If lastCol >= 1 And lastRow > startRow Then
                Set dataRng = ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, lastCol))
                MultiSheetRawDataJSON = RangeRawDataJSON(dataRng)
                Exit Function
            End If
        End If
    Next i
End Function

Private Function RangeRowJSON(ByVal rowRange As Range) As String
    Dim values() As String
    Dim c As Long
    ReDim values(1 To rowRange.Columns.Count)

    For c = 1 To rowRange.Columns.Count
        values(c) = JsonValue(rowRange.Cells(1, c).Value)
    Next c

    RangeRowJSON = "[" & Join(values, ",") & "]"
End Function

Private Function BuildDashboardConfigJSON(ByVal dashTitle As String, _
                                          ByVal theme As String, _
                                          ByVal currencyCode As String, _
                                          ByVal smartLock As Boolean, _
                                          ByVal gridlinesShow As Boolean, _
                                          ByVal rawDataJson As String) As String
    Dim headers() As String
    headers = HeadersFromRawData(rawDataJson)

    Dim firstColumn As String
    Dim numericColumn As String
    Dim dateColumn As String
    firstColumn = FirstHeader(headers)
    numericColumn = FindNumericHeader(headers)
    dateColumn = FindDateHeader(headers)

    If Len(dateColumn) = 0 Then dateColumn = firstColumn
    If Len(numericColumn) = 0 Then numericColumn = firstColumn

    Dim config As String
    config = "{" & _
             JsonPropertyName("schemaVersion") & CStr(DASHBOARD_SCHEMA_VERSION) & "," & _
             JsonPropertyName("application") & JsonQuote(DASHBOARD_APPLICATION_NAME) & "," & _
             JsonPropertyName("title") & JsonQuote(dashTitle) & "," & _
             JsonPropertyName("theme") & JsonQuote(TemplateThemeName(theme)) & "," & _
             JsonPropertyName("currency") & JsonQuote(currencyCode) & "," & _
             JsonPropertyName("showGrid") & JsonBoolean(gridlinesShow) & "," & _
             JsonPropertyName("smartLock") & JsonBoolean(smartLock) & "," & _
             JsonPropertyName("visuals") & "[]," & _
             JsonPropertyName("layout") & "{""columns"":12,""rowHeight"":80,""margin"":8}," & _
             JsonPropertyName("data") & "{" & _
             JsonPropertyName("sourceType") & JsonQuote("excel-range") & "," & _
             JsonPropertyName("headers") & HeaderArrayJSON(headers) & "}," & _
             JsonPropertyName("interactionState") & "{""filters"":{},""zoom"":{}}," & _
             JsonPropertyName("kpis") & "[]," & _
             JsonPropertyName("charts") & "[]," & _
             JsonPropertyName("slicers") & "[]," & _
             JsonPropertyName("logoImage") & JsonQuote("") & "," & _
             JsonPropertyName("logoIcon") & JsonQuote("") & "}"

    BuildDashboardConfigJSON = config
End Function

Private Function KpiJSON(ByVal title As String, ByVal columnName As String, _
                         ByVal aggregation As String, ByVal iconName As String, _
                         ByVal x As Long, ByVal y As Long) As String
    KpiJSON = "{" & _
              JsonPropertyName("title") & JsonQuote(title) & _
              "," & JsonPropertyName("col") & JsonQuote(columnName) & _
              "," & JsonPropertyName("agg") & JsonQuote(aggregation) & _
              "," & JsonPropertyName("icon") & JsonQuote(iconName) & _
              "," & JsonPropertyName("image") & JsonQuote("") & _
              "," & JsonPropertyName("x") & CStr(x) & _
              "," & JsonPropertyName("y") & CStr(y) & _
              "," & JsonPropertyName("w") & "3," & _
              JsonPropertyName("h") & "2}"
End Function

Private Function ChartJSON(ByVal title As String, ByVal chartType As String, _
                           ByVal xColumn As String, ByVal yColumn As String, _
                           ByVal aggregation As String, ByVal hierarchy As String, _
                           ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As String
    ChartJSON = "{" & _
                JsonPropertyName("title") & JsonQuote(title) & _
                "," & JsonPropertyName("type") & JsonQuote(chartType) & _
                "," & JsonPropertyName("top") & "0," & _
                JsonPropertyName("xCol") & JsonQuote(xColumn) & _
                "," & JsonPropertyName("yCol") & JsonQuote(yColumn) & _
                "," & JsonPropertyName("agg") & JsonQuote(aggregation) & _
                "," & JsonPropertyName("xLabel") & JsonQuote("") & _
                "," & JsonPropertyName("yLabel") & JsonQuote("") & _
                "," & JsonPropertyName("hierarchy") & JsonQuote(hierarchy) & _
                "," & JsonPropertyName("labelMode") & JsonQuote("auto") & _
                "," & JsonPropertyName("labelsVisible") & "true," & _
                JsonPropertyName("x") & CStr(x) & _
                "," & JsonPropertyName("y") & CStr(y) & _
                "," & JsonPropertyName("w") & CStr(width) & _
                "," & JsonPropertyName("h") & CStr(height) & "}"
End Function

Private Function JsonPropertyName(ByVal name As String) As String
    JsonPropertyName = JsonQuote(name) & ":"
End Function

Private Function HeadersFromRawData(ByVal rawDataJson As String) As String()
    Dim headerJson As String
    Dim headers() As String
    Dim position As Long
    Dim itemCount As Long
    Dim headerValue As String

    position = InStr(rawDataJson, "[[")
    If position = 0 Then Exit Function
    position = position + 2

    Do
        If position > Len(rawDataJson) Or Mid$(rawDataJson, position, 1) = "]" Then Exit Do
        headerValue = ReadJsonString(rawDataJson, position)
        If Len(headerValue) = 0 And Mid$(rawDataJson, position, 1) <> """" Then Exit Do

        If itemCount = 0 Then
            ReDim headers(0 To 0)
        Else
            ReDim Preserve headers(0 To itemCount)
        End If
        headers(itemCount) = headerValue
        itemCount = itemCount + 1

        Do While position <= Len(rawDataJson) And Mid$(rawDataJson, position, 1) <> "," And Mid$(rawDataJson, position, 1) <> "]"
            position = position + 1
        Loop
        If Mid$(rawDataJson, position, 1) = "," Then position = position + 1
    Loop

    HeadersFromRawData = headers
End Function

Private Function ReadJsonString(ByVal text As String, ByRef position As Long) As String
    Dim result As String
    Dim ch As String
    Dim escaped As Boolean

    Do While position <= Len(text) And (Mid$(text, position, 1) = " " Or Mid$(text, position, 1) = vbTab)
        position = position + 1
    Loop
    If Mid$(text, position, 1) <> """" Then Exit Function

    position = position + 1
    Do While position <= Len(text)
        ch = Mid$(text, position, 1)
        position = position + 1

        If escaped Then
            Select Case ch
                Case """": result = result & """"
                Case "\": result = result & "\"
                Case "n": result = result & vbLf
                Case "r": result = result & vbCr
                Case "t": result = result & vbTab
                Case Else: result = result & ch
            End Select
            escaped = False
        ElseIf ch = "\" Then
            escaped = True
        ElseIf ch = """" Then
            Exit Do
        Else
            result = result & ch
        End If
    Loop

    ReadJsonString = result
End Function

Private Function HeaderArrayJSON(ByRef headers() As String) As String
    Dim i As Long
    Dim result As String
    On Error GoTo NoHeaders

    result = "["
    For i = LBound(headers) To UBound(headers)
        If i > LBound(headers) Then result = result & ","
        result = result & JsonQuote(headers(i))
    Next i
    HeaderArrayJSON = result & "]"
    Exit Function

NoHeaders:
    HeaderArrayJSON = "[]"
End Function

Private Function FirstHeader(ByRef headers() As String) As String
    On Error GoTo NoHeader
    FirstHeader = headers(LBound(headers))
    Exit Function
NoHeader:
    FirstHeader = "Value"
End Function

Private Function FindNumericHeader(ByRef headers() As String) As String
    Dim i As Long
    Dim h As String
    On Error GoTo NoNumeric

    For i = LBound(headers) To UBound(headers)
        h = LCase$(headers(i))
        If InStr(h, "amount") > 0 Or InStr(h, "sales") > 0 Or InStr(h, "revenue") > 0 Or _
           InStr(h, "price") > 0 Or InStr(h, "total") > 0 Or InStr(h, "cost") > 0 Or _
           InStr(h, "profit") > 0 Or InStr(h, "value") > 0 Or InStr(h, "quantity") > 0 Or _
           InStr(h, "qty") > 0 Or InStr(h, "count") > 0 Then
            FindNumericHeader = headers(i)
            Exit Function
        End If
    Next i

    FindNumericHeader = FirstHeader(headers)
    Exit Function
NoNumeric:
    FindNumericHeader = "Value"
End Function

Private Function FindDateHeader(ByRef headers() As String) As String
    Dim i As Long
    Dim h As String
    On Error GoTo NoDate

    For i = LBound(headers) To UBound(headers)
        h = LCase$(headers(i))
        If InStr(h, "date") > 0 Or InStr(h, "month") > 0 Or InStr(h, "year") > 0 Or InStr(h, "time") > 0 Then
            FindDateHeader = headers(i)
            Exit Function
        End If
    Next i

NoDate:
    FindDateHeader = ""
End Function

Private Function JsonUnquote(ByVal text As String) As String
    text = Trim$(text)
    If Left$(text, 1) = """" Then text = Mid$(text, 2)
    If Right$(text, 1) = """" Then text = Left$(text, Len(text) - 1)
    text = Replace(text, Chr$(92) & Chr$(34), Chr$(34))
    text = Replace(text, Chr$(92) & Chr$(92), Chr$(92))
    JsonUnquote = text
End Function

Private Function TemplateThemeName(ByVal theme As String) As String
    Select Case LCase$(Trim$(theme))
        Case "ocean", "forest", "rose": TemplateThemeName = "Ocean"
        Case "midnight", "navy", "slate": TemplateThemeName = "Midnight"
        Case Else: TemplateThemeName = DASHBOARD_DEFAULT_THEME
    End Select
End Function

Private Function LoadDashboardTemplate() As String
    Dim candidate As String

    candidate = ThisWorkbook.Path & "\" & TEMPLATE_FILE_NAME
    If Dir$(candidate) <> "" Then
        LoadDashboardTemplate = ReadUtf8(candidate)
        If Len(LoadDashboardTemplate) > 0 Then Exit Function
    End If

    candidate = ThisWorkbook.Path & "\" & TEMPLATE_ASSET_FOLDER & "\" & TEMPLATE_FILE_NAME
    If Dir$(candidate) <> "" Then
        LoadDashboardTemplate = ReadUtf8(candidate)
        If Len(LoadDashboardTemplate) > 0 Then Exit Function
    End If

    candidate = Environ$("APPDATA") & "\Microsoft\AddIns\" & TEMPLATE_FILE_NAME
    If Dir$(candidate) <> "" Then
        LoadDashboardTemplate = ReadUtf8(candidate)
        If Len(LoadDashboardTemplate) > 0 Then Exit Function
    End If

    ' The template is hosted separately because it is too large for a VBA module.
    ' Cache it beside the installed add-in so future dashboard creation is offline-safe.
    If DownloadDashboardTemplate(candidate) Then
        LoadDashboardTemplate = ReadUtf8(candidate)
    End If
End Function

Private Function DownloadDashboardTemplate(ByVal destination As String) As Boolean
    On Error GoTo failSafe
    Dim http As Object
    Dim stream As Object

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", TEMPLATE_URL, False
    http.Send
    If http.Status <> 200 Then Exit Function

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile destination, 2
    stream.Close
    DownloadDashboardTemplate = (Dir$(destination) <> "")
    Exit Function

failSafe:
    DownloadDashboardTemplate = False
End Function

Private Function ReplaceTemplateSection(ByVal html As String, ByVal startMarker As String, _
                                        ByVal endMarker As String, ByVal replacement As String) As String
    Dim startAt As Long
    Dim endAt As Long

    startAt = InStr(1, html, startMarker, vbBinaryCompare)
    endAt = InStr(1, html, endMarker, vbBinaryCompare)
    If startAt = 0 Or endAt = 0 Or endAt <= startAt Then
        Err.Raise vbObjectError + 106, , "Power Dashboard template placeholders are invalid."
    End If

    ReplaceTemplateSection = Left$(html, startAt + Len(startMarker) - 1) & vbLf & _
                             replacement & vbLf & Mid$(html, endAt)
End Function

Private Function ApplyDashboardSettings(ByVal html As String, ByVal dashTitle As String, _
                                        ByVal theme As String, ByVal currencyCode As String, _
                                        ByVal smartLock As Boolean, ByVal gridlinesShow As Boolean) As String
    Dim titleScript As String
    titleScript = "document.getElementById('mainTitle').innerText = " & JsonQuote(dashTitle) & ";" & _
                  "currencySym = " & JsonQuote(CurrencySymbol(currencyCode)) & ";" & _
                  "showGrid = " & JsonBoolean(gridlinesShow) & ";" & _
                  "window.XLSmartSmartLock = " & JsonBoolean(smartLock) & ";" & _
                  "window.XLSmartGridlinesShow = " & JsonBoolean(gridlinesShow) & ";"

    html = Replace(html, "document.getElementById('mainTitle').innerText = 'XLSmart Dashboard';", titleScript)
    ApplyDashboardSettings = html
End Function

Private Function CurrencySymbol(ByVal currencyCode As String) As String
    Select Case UCase$(Trim$(currencyCode))
        Case "BDT": CurrencySymbol = ChrW$(&H9F3)
        Case "INR": CurrencySymbol = ChrW$(&H20B9)
        Case "EUR": CurrencySymbol = ChrW$(&H20AC)
        Case "USD": CurrencySymbol = "$"
        Case Else: CurrencySymbol = ""
    End Select
End Function

Private Function JsonBoolean(ByVal value As Boolean) As String
    If value Then
        JsonBoolean = "true"
    Else
        JsonBoolean = "false"
    End If
End Function

Private Function JsonValue(ByVal value As Variant) As String
    Select Case VarType(value)
        Case vbEmpty, vbNull, vbError
            JsonValue = "null"
        Case vbDate
            If CDbl(value) = Int(CDbl(value)) Then
                JsonValue = JsonQuote(Format$(CDate(value), "yyyy-mm-dd"))
            Else
                JsonValue = JsonQuote(Format$(CDate(value), "yyyy-mm-dd hh:nn:ss"))
            End If
        Case vbBoolean
            JsonValue = JsonBoolean(CBool(value))
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            JsonValue = Replace(Trim$(Str$(CDbl(value))), ",", ".")
        Case Else
            JsonValue = JsonQuote(CStr(value))
    End Select
End Function

Private Function JsonQuote(ByVal text As String) As String
    text = Replace(text, Chr$(92), Chr$(92) & Chr$(92))
    text = Replace(text, Chr$(34), Chr$(92) & Chr$(34))
    text = Replace(text, vbCrLf, "\n")
    text = Replace(text, vbCr, "\n")
    text = Replace(text, vbLf, "\n")
    text = Replace(text, vbTab, "\t")
    JsonQuote = Chr$(34) & text & Chr$(34)
End Function

Private Function ReadUtf8(ByVal filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadUtf8 = stream.ReadText
    stream.Close
End Function

Private Sub WriteUtf8(ByVal filePath As String, ByVal text As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile filePath, 2
    stream.Close
End Sub

Private Sub OpenDashboard(ByVal htmlPath As String)
    On Error Resume Next
    ThisWorkbook.FollowHyperlink htmlPath
    If Err.Number <> 0 Then
        Err.Clear
        Shell "cmd /c start """" """ & htmlPath & """", vbHide
    End If
    On Error GoTo 0
End Sub

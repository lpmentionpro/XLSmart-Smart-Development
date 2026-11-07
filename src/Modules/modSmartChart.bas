Attribute VB_Name = "modSmartChart"
Option Explicit

Public Sub CreateWebChart( _
    ByVal rng As Range, _
    ByVal chartType As String, _
    ByVal chartTitle As String, _
    ByVal themeName As String, _
    ByVal xAxisTitle As String, _
    ByVal yAxisTitle As String)

    Dim html As String
    Dim savePath As String
    Dim ff As Integer

    html = BuildChartHTML( _
        rng, _
        chartType, _
        chartTitle, _
        themeName, _
        xAxisTitle, _
        yAxisTitle)

    savePath = Environ$("TEMP") & "\smart_chart.html"

    ff = FreeFile

    Open savePath For Output As #ff
    Print #ff, html
    Close #ff

    ThisWorkbook.FollowHyperlink savePath

End Sub


Private Function BuildChartHTML( _
    ByVal rng As Range, _
    ByVal chartType As String, _
    ByVal chartTitle As String, _
    ByVal themeName As String, _
    ByVal xAxisTitle As String, _
    ByVal yAxisTitle As String) As String

    Dim labels As String
    Dim values As String
    Dim i As Long

    Dim bgColor As String
    Dim textColor As String
    Dim chartColor As String

    Dim h As String
    Dim isHorizontal As String

    ' data build
    For i = 2 To rng.Rows.count

        labels = labels & "'" & _
                 Replace(rng.Cells(i, 1).Value, "'", "") & "',"

        values = values & _
                 val(rng.Cells(i, 2).Value) & ","

    Next i

    If labels <> "" Then
        labels = left(labels, Len(labels) - 1)
    End If

    If values <> "" Then
        values = left(values, Len(values) - 1)
    End If


    ' theme
    Select Case themeName

        Case "Dark"

            bgColor = "#111827"
            textColor = "#FFFFFF"
            chartColor = "#00D4FF"

        Case "Ocean"

            bgColor = "#EAF6FF"
            textColor = "#0F172A"
            chartColor = "#0077B6"

        Case "Midnight"

            bgColor = "#0F172A"
            textColor = "#FFFFFF"
            chartColor = "#8B5CF6"

        Case "Corporate"

            bgColor = "#F3F4F6"
            textColor = "#111827"
            chartColor = "#005B96"

        Case Else

            bgColor = "#FFFFFF"
            textColor = "#111827"
            chartColor = "#2563EB"

    End Select


    ' chart map
    Select Case chartType

    Case "Column Chart"
        chartType = "bar"
        isHorizontal = "false"

    Case "Bar Chart"
        chartType = "bar"
        isHorizontal = "true"

    Case "Line Chart"
        chartType = "line"
        isHorizontal = "false"

    Case "Area Chart"
        chartType = "line"
        isHorizontal = "false"

    Case "Pie Chart"
        chartType = "pie"
        isHorizontal = "false"

    Case "Doughnut Chart"
        chartType = "pie"
        isHorizontal = "false"

    Case "Scatter Chart"
        chartType = "scatter"
        isHorizontal = "false"

    Case "Radar Chart"
        chartType = "radar"
        isHorizontal = "false"

    Case Else
        chartType = "bar"
        isHorizontal = "false"

End Select


    ' html start
    h = "<!doctype html><html><head>"
    h = h & "<meta charset='utf-8'>"
    h = h & "<meta name='viewport' content='width=device-width,initial-scale=1'>"

    h = h & "<script src='https://cdn.jsdelivr.net/npm/echarts/dist/echarts.min.js'></script>"

    h = h & "<title>" & chartTitle & "</title>"

    h = h & "<style>"
    h = h & "body{margin:0;background:" & bgColor & ";font-family:Segoe UI;overflow:hidden;}"

    h = h & ".top{padding:20px;color:" & textColor & ";}"
    h = h & ".top h1{margin:0;font-size:32px;font-weight:700;}"

    h = h & ".toolbar{padding:0 20px 15px;}"

    h = h & "button{background:" & chartColor & ";border:none;padding:10px 18px;"
    h = h & "border-radius:12px;color:#fff;cursor:pointer;margin-right:10px;font-weight:600;}"

    h = h & "#chart{width:97%;height:92vh;margin:20px auto;"
    h = h & "background:#fff;border-radius:22px;box-shadow:0 8px 30px rgba(0,0,0,.12);}"

    h = h & "</style>"
    h = h & "</head><body>"

    h = h & "<div id='chart'></div>"

    h = h & "<script>"

h = h & "var chart = echarts.init(document.getElementById('chart'));"

h = h & "var option = {"

' animation
h = h & "animation:true,"
h = h & "animationDuration:1800,"
h = h & "animationEasing:'cubicOut',"

' title
h = h & "title:{"
h = h & "text:'" & chartTitle & "',"
h = h & "left:'center',"
h = h & "top:22,"
h = h & "textStyle:{"
h = h & "fontSize:34,"
h = h & "fontWeight:'700',"
h = h & "color:'" & textColor & "'"
h = h & "}"
h = h & "},"

' tooltip premium
h = h & "tooltip:{"
h = h & "trigger:'axis',"
h = h & "backgroundColor:'rgba(20,20,20,.92)',"
h = h & "borderRadius:18,"
h = h & "padding:[14,18],"
h = h & "textStyle:{fontSize:13,color:'#fff'}"
h = h & "},"

' right icons only
h = h & "toolbox:{"
h = h & "right:28,"
h = h & "top:22,"
h = h & "itemGap:12,"
h = h & "iconStyle:{borderColor:'#6C7BFF'},"
h = h & "feature:{"
h = h & "saveAsImage:{},"
h = h & "restore:{},"
h = h & "dataView:{show:false},"
h = h & "dataZoom:{}"
h = h & "}"
h = h & "},"

' grid
h = h & "grid:{"
h = h & "left:'6%',"
h = h & "right:'5%',"
h = h & "bottom:'12%',"
h = h & "top:'16%',"
h = h & "containLabel:true"
h = h & "},"

' x axis
h = h & "xAxis:{"

If isHorizontal = "true" Then
    h = h & "type:'value',"
Else
    h = h & "type:'category',"
    h = h & "data:[" & labels & "],"
End If
h = h & "name:'" & xAxisTitle & "',"
h = h & "nameLocation:'middle',"
h = h & "nameGap:50,"
h = h & "data:[" & labels & "],"
h = h & "axisLabel:{"
h = h & "color:'" & textColor & "',"
h = h & "fontSize:13,"
h = h & "fontWeight:'600'"
h = h & "},"
h = h & "nameTextStyle:{"
h = h & "color:'" & textColor & "',"
h = h & "fontWeight:'bold',"
h = h & "fontSize:16"
h = h & "},"
h = h & "axisLine:{"
h = h & "lineStyle:{color:'#A1A1AA'}"
h = h & "}"
h = h & "},"

h = h & "yAxis:{"

If isHorizontal = "true" Then
    h = h & "type:'category',"
    h = h & "data:[" & labels & "],"
Else
    h = h & "type:'value',"
End If
h = h & "name:'" & yAxisTitle & "',"
h = h & "nameTextStyle:{"
h = h & "color:'" & textColor & "',"
h = h & "fontWeight:'bold',"
h = h & "fontSize:16"
h = h & "},"
h = h & "axisLabel:{"
h = h & "color:'" & textColor & "'"
h = h & "},"
h = h & "splitLine:{"
h = h & "lineStyle:{"
h = h & "type:'dashed',"
h = h & "opacity:.25"
h = h & "}"
h = h & "}"
h = h & "},"

' series
h = h & "series:[{"

h = h & "data:[" & values & "],"
h = h & "type:'" & chartType & "',"

' line chart smooth
h = h & "smooth:true,"
h = h & "symbol:'circle',"
h = h & "symbolSize:8,"

' bar width
h = h & "barWidth:'56%',"

' line style
h = h & "lineStyle:{width:4},"

' area vibe for line chart
h = h & "areaStyle:{"
h = h & "opacity:.12"
h = h & "},"

' gradient color
h = h & "itemStyle:{"
h = h & "borderRadius:[16,16,0,0],"
h = h & "color:new echarts.graphic.LinearGradient("
h = h & "0,0,0,1,["
h = h & "{offset:0,color:'#4F8CFF'},"
h = h & "{offset:1,color:'#7C4DFF'}"
h = h & "])"
h = h & "},"

' hover
h = h & "emphasis:{"
h = h & "focus:'series',"
h = h & "scale:true"
h = h & "},"

' labels
h = h & "label:{"
h = h & "show:true,"
h = h & "position:'top',"
h = h & "fontWeight:'bold',"
h = h & "fontSize:12"
h = h & "}"

h = h & "}]"

h = h & "};"

h = h & "chart.setOption(option);"

h = h & "window.addEventListener('resize', function(){"
h = h & "chart.resize();"
h = h & "});"

h = h & "</script>"

    h = h & "</body></html>"

    BuildChartHTML = h

End Function

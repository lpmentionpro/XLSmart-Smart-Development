Option Explicit

                           
Public Sub CreateDashboard(ByVal dashTitle As String, _
                           ByVal theme As String, _
                           ByVal currencyCodeCode As String, _
                           ByVal isMulti As Boolean, _
                           ByVal rangeOrSheets As String, _
                           ByVal startRow As Long, _
                           ByVal smartLock As Boolean)
    On Error GoTo CleanFail

    Application.StatusBar = "Power Dashboard: reading data..."

    Dim dataWb As Workbook
    Dim jsonData As String

    If isMulti Then

        Set dataWb = ActiveWorkbook

        If dataWb Is Nothing Then
            Err.Raise vbObjectError + 100, , _
                "No active workbook found."
        End If

        If dataWb Is ThisWorkbook Then
            Err.Raise vbObjectError + 101, , _
                "Please activate the workbook that contains your data, then run Power Dashboard again."
        End If

        jsonData = MultiSheetJSON(dataWb, rangeOrSheets, startRow)

    Else

        Dim rng As Range

        Set rng = ResolveRange(rangeOrSheets)

        If rng Is Nothing Then
            Err.Raise vbObjectError + 102, , _
                "Invalid data range: " & rangeOrSheets
        End If

        Set dataWb = rng.Worksheet.Parent
        jsonData = SingleRangeJSON(rng)

    End If

    If Len(jsonData) = 0 Or jsonData = "{}" Then
        Err.Raise vbObjectError + 103, , _
            "No readable data found."
    End If

    Application.StatusBar = _
    "Power Dashboard: creating HTML..."

Dim htmlContent As String
Dim htmlPath As String
Dim ff As Integer

htmlContent = BuildDashboardHTML( _
                    jsonData, _
                    dashTitle, _
                    theme, _
                    currencyCodeCode, _
                    smartLock)

htmlPath = Environ$("TEMP") & "\power_dashboard.html"

ff = FreeFile

Open htmlPath For Output As #ff
Print #ff, htmlContent
Close #ff

If Dir(htmlPath) = "" Then
    Err.Raise vbObjectError + 104, , _
        "Dashboard HTML file could not be created."
End If

Application.StatusBar = "Opening dashboard..."

On Error Resume Next

ThisWorkbook.FollowHyperlink htmlPath

If Err.Number <> 0 Then
    Err.Clear
    Shell "cmd /c start """" """ & htmlPath & """", vbHide
End If

On Error GoTo CleanFail

Application.StatusBar = False

MsgBox "Dashboard created successfully!", vbInformation

Exit Sub
CleanFail:

    Application.StatusBar = False

    MsgBox Err.Description, vbCritical, _
        "Power Dashboard"

End Sub

Private Function ResolveRange(ByVal addressText As String) As Range

    On Error GoTo failSafe

    Dim parts() As String
    Dim wb As Workbook
    Dim ws As Worksheet

    parts = Split(addressText, "|")

    If UBound(parts) <> 2 Then
        Exit Function
    End If

    Set wb = Nothing

    On Error Resume Next
    Set wb = Workbooks(parts(0))
    On Error GoTo failSafe

    If wb Is Nothing Then
        Exit Function
    End If

    Set ws = wb.Worksheets(parts(1))

    If ws Is Nothing Then
        Exit Function
    End If

    Set ResolveRange = ws.Range(parts(2))
    Exit Function

failSafe:
    Set ResolveRange = Nothing

End Function

Private Function SingleRangeJSON(ByVal rng As Range) As String
    If rng Is Nothing Then Exit Function
    If rng.Rows.count < 2 Or rng.Columns.count < 1 Then Exit Function

    Dim headers() As String
    headers = HeadersFromRange(rng.Rows(1))

    Dim rowsJson As String
    rowsJson = RowsFromRange(rng.Offset(1, 0).Resize(rng.Rows.count - 1), headers)

    SingleRangeJSON = "{" & JsonQuote(rng.Worksheet.Name) & _
                      ":{" & JsonQuote("headers") & ":[" & JoinQuoted(headers) & "]," & _
                      JsonQuote("rows") & ":[" & rowsJson & "]}}"
End Function

Private Function MultiSheetJSON(ByVal wb As Workbook, ByVal sheetList As String, ByVal startRow As Long) As String
    Dim parts() As String
    parts = Split(sheetList, ",")

    Dim result As String
    Dim firstSheet As Boolean
    firstSheet = True
    result = "{"

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim sheetIndex As Long
        sheetIndex = val(Trim$(parts(i)))
        If sheetIndex >= 1 And sheetIndex <= wb.Worksheets.count Then
            Dim ws As Worksheet
            Set ws = wb.Worksheets(sheetIndex)

            Dim lastCol As Long
            Dim lastRow As Long
            lastCol = ws.Cells(startRow, ws.Columns.count).End(xlToLeft).Column
            lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

            If lastCol >= 1 And lastRow > startRow Then
                Dim dataRng As Range
                Set dataRng = ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, lastCol))

                Dim headers() As String
                headers = HeadersFromRange(dataRng.Rows(1))

                If Not firstSheet Then result = result & ","
                firstSheet = False
                result = result & JsonQuote(ws.Name) & _
                         ":{" & JsonQuote("headers") & ":[" & JoinQuoted(headers) & "]," & _
                         JsonQuote("rows") & ":[" & _
                         RowsFromRange(dataRng.Offset(1, 0).Resize(dataRng.Rows.count - 1), headers) & "]}"
            End If
        End If
    Next i

    MultiSheetJSON = result & "}"
End Function

Private Function HeadersFromRange(ByVal headerRow As Range) As String()
    Dim headers() As String
    ReDim headers(1 To headerRow.Columns.count)

    Dim c As Long
    For c = 1 To headerRow.Columns.count
        headers(c) = Trim$(CStr(headerRow.Cells(1, c).Value))
        If Len(headers(c)) = 0 Then headers(c) = "Column" & c
    Next c

    HeadersFromRange = headers
End Function

Private Function RowsFromRange(ByVal bodyRng As Range, ByRef headers() As String) As String
    Dim rowParts() As String
    ReDim rowParts(1 To bodyRng.Rows.count)

    Dim r As Long
    Dim c As Long
    For r = 1 To bodyRng.Rows.count
        Dim obj As String
        obj = "{"
        For c = 1 To bodyRng.Columns.count
            If c > 1 Then obj = obj & ","
            obj = obj & JsonQuote(headers(c)) & ":" & JsonValue(bodyRng.Cells(r, c).Value)
        Next c
        obj = obj & "}"
        rowParts(r) = obj
    Next r

    RowsFromRange = Join(rowParts, ",")
End Function

Private Function JoinQuoted(ByRef values() As String) As String
    Dim parts() As String
    ReDim parts(LBound(values) To UBound(values))

    Dim i As Long
    For i = LBound(values) To UBound(values)
        parts(i) = JsonQuote(values(i))
    Next i

    JoinQuoted = Join(parts, ",")
End Function

Private Function JsonValue(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then
        JsonValue = "null"
    ElseIf IsDate(v) And Not IsNumeric(v) Then
        JsonValue = JsonQuote(Format$(CDate(v), "yyyy-mm-dd"))
    ElseIf IsNumeric(v) Then
        JsonValue = Trim$(Str$(CDbl(v)))
    Else
        JsonValue = JsonQuote(CStr(v))
    End If
End Function

Private Function JsonQuote(ByVal text As String) As String
    text = Replace(text, Chr$(92), Chr$(92) & Chr$(92))
    text = Replace(text, Chr$(34), Chr$(92) & Chr$(34))
    text = Replace(text, vbCrLf, "\n")
    text = Replace(text, vbCr, "\n")
    text = Replace(text, vbLf, "\n")
    JsonQuote = """" & text & """"
End Function

Private Function BuildDashboardHTML(ByVal jsonData As String, _
                                    ByVal dashTitle As String, _
                                    ByVal theme As String, _
                                    ByVal currencyCode As String, _
                                    ByVal smartLock As Boolean) As String
    Dim bg As String, surface As String, textColor As String, accent As String
    Select Case LCase$(theme)
        Case "ocean": bg = "#f3f7ff": surface = "#ffffff": textColor = "#172033": accent = "#2563eb"
        Case "forest": bg = "#f0fdf4": surface = "#ffffff": textColor = "#14351f": accent = "#16a34a"
        Case "rose": bg = "#fff1f2": surface = "#ffffff": textColor = "#4c0519": accent = "#e11d48"
        Case "navy": bg = "#0c1a2e": surface = "#162035": textColor = "#e7eefc": accent = "#60a5fa"
        Case "slate": bg = "#0f172a": surface = "#1e293b": textColor = "#e2e8f0": accent = "#38bdf8"
        Case Else: bg = "#0c1120": surface = "#111827": textColor = "#e2e8f0": accent = "#818cf8"
    End Select

    Dim lockText As String
    lockText = IIf(smartLock, "true", "false")

    Dim h As String
    h = "<!doctype html><html><head><meta charset='utf-8'>"
    h = h & "<meta name='viewport' content='width=device-width,initial-scale=1'>"
    h = h & "<title>" & HtmlEncode(dashTitle) & "</title>"
    h = h & "<style>"
    h = h & ":root{--bg:" & bg & ";--panel:" & surface & ";--txt:" & textColor & ";--acc:" & accent & ";--line:rgba(148,163,184,.3);--mut:#64748b;--danger:#ef4444}"
    h = h & "*{box-sizing:border-box}body{margin:0;background:linear-gradient(135deg,var(--bg),#eef6ff);color:var(--txt);font-family:Segoe UI,Arial,sans-serif;min-height:100vh}"
    h = h & "body.dark{background:linear-gradient(135deg,#0f172a,#1e293b)}header{position:sticky;top:0;z-index:20;background:rgba(255,255,255,.92);backdrop-filter:blur(12px);border-bottom:1px solid var(--line);box-shadow:0 10px 28px rgba(15,23,42,.1)}"
    h = h & ".top{height:64px;display:flex;align-items:center;gap:10px;padding:0 16px;overflow:auto}.brand{display:flex;align-items:center;gap:10px;margin-right:auto;min-width:250px}.logo{width:38px;height:38px;border-radius:10px;border:1px dashed #a5b4fc;background:#f8fafc;display:grid;place-items:center;cursor:pointer;overflow:hidden}.logo img{width:100%;height:100%;object-fit:cover}h1{margin:0;font-size:20px;color:#2563eb;white-space:nowrap}"
    h = h & "button,select,input,textarea{font:inherit;border:1px solid var(--line);border-radius:9px;padding:9px 11px;background:#fff;color:#111827}button{cursor:pointer;font-weight:800;box-shadow:0 6px 14px rgba(15,23,42,.08)}button:hover{transform:translateY(-1px)}.primary{background:var(--acc);border-color:var(--acc);color:#fff}.green{background:#16a34a;color:#fff}.rose{background:#e11d48;color:#fff}.amber{background:#d97706;color:#fff}.darkbtn{background:#111827;color:#fff}.danger{background:var(--danger);color:#fff}"
    h = h & ".wrap{padding:18px 22px 78px}.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;align-items:stretch}.slot{min-height:124px;border:1.8px dashed #d8dee9;border-radius:14px;background:rgba(255,255,255,.68);display:flex;align-items:center;justify-content:center;color:#64748b;font-weight:800;position:relative;transition:.16s}.slot.big{min-height:330px;grid-column:span 2}.slot.wide{grid-column:span 4}.slot:hover{border-color:var(--acc);background:#fff}.slot.dragOver{outline:3px solid rgba(99,102,241,.25)}"
    h = h & ".widget{width:100%;height:100%;min-height:inherit;background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:16px;box-shadow:0 14px 34px rgba(15,23,42,.1);resize:both;overflow:auto;position:relative}.widget[draggable=true]{cursor:grab}.tools{position:absolute;right:10px;top:8px;display:flex;gap:5px}.tools button{padding:5px 7px;border-radius:7px;box-shadow:none}.title{font-size:16px;font-weight:900;margin:0 0 12px}.k{font-size:11px;text-transform:uppercase;color:var(--mut);font-weight:900}.v{font-size:30px;font-weight:900;margin-top:8px}"
    h = h & ".kpi:before{content:'';position:absolute;left:0;top:0;bottom:0;width:5px;background:linear-gradient(180deg,var(--acc),#22c55e)}"
    h = h & ".slicerBtns{display:flex;gap:8px;flex-wrap:wrap}.pill{border-radius:999px;background:#fff;border:1px solid var(--line);box-shadow:none}.pill.on{background:var(--acc);color:#fff}.chartBox{height:285px}.chartBox svg{width:100%;height:100%;display:block}.axis{stroke:#cbd5e1}.bar{fill:var(--acc)}.line{fill:none;stroke:var(--acc);stroke-width:3}.label{font-size:11px;fill:#334155;font-weight:800}.dark .label{fill:#e2e8f0}.note textarea{width:100%;min-height:90px;border:0;background:#f8fafc;resize:vertical}.noteTitle{font-weight:900;width:100%;border:0;background:transparent;padding:3px 0 9px}"
    h = h & ".tableWrap{max-height:360px;overflow:auto;border:1px solid var(--line);border-radius:10px;background:#fff}table{width:100%;border-collapse:separate;border-spacing:0;font-size:13px;color:#111827}th,td{padding:10px 11px;border-bottom:1px solid #e5e7eb;text-align:left;white-space:nowrap}th{position:sticky;top:0;background:#f8fafc;color:#475569;font-size:11px;text-transform:uppercase;z-index:2}tbody tr:nth-child(even){background:#fbfdff}tbody tr:hover{background:#eef6ff}.small{font-size:12px;color:var(--mut);font-weight:800}"
    h = h & ".modal{position:fixed;inset:0;background:rgba(15,23,42,.48);display:none;align-items:center;justify-content:center;z-index:50}.modal.show{display:flex}.box{width:min(880px,94vw);max-height:88vh;overflow:auto;background:#fff;color:#111827;border-radius:18px;padding:24px;box-shadow:0 30px 90px rgba(15,23,42,.35)}.box h2{margin:0 0 16px}.section{border-top:1px solid #e5e7eb;padding-top:14px;margin-top:14px}.form{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}.field{display:flex;flex-direction:column;gap:6px}.field label{font-size:11px;font-weight:900;color:#64748b;text-transform:uppercase}"
    h = h & ".actions{display:flex;justify-content:flex-end;gap:10px;margin-top:18px}.templates{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}.tpl{border:1px solid #e5e7eb;border-radius:14px;padding:14px;text-align:center;cursor:pointer}.tpl:hover{border-color:var(--acc);box-shadow:0 12px 28px rgba(15,23,42,.1)}.sk{height:70px;display:grid;grid-template-columns:repeat(4,1fr);gap:5px;background:#f8fafc;border-radius:10px;padding:8px}.sk i{background:#a78bfa;border-radius:5px}"
    h = h & ".footer{position:fixed;left:0;right:0;bottom:0;height:54px;background:rgba(255,255,255,.94);border-top:1px solid var(--line);display:flex;align-items:center;gap:12px;padding:0 20px;z-index:19}.footer b{color:#2563eb;margin-right:auto}@media(max-width:950px){.grid{grid-template-columns:1fr}.slot.big,.slot.wide{grid-column:span 1}.form,.templates{grid-template-columns:1fr}.top{height:auto;padding:10px}.brand{width:100%}}@media print{header,.footer,.tools,.modal{display:none!important}body{background:#fff}.wrap{padding:0}.grid{display:block}.slot{border:0;page-break-inside:avoid;margin:0 0 14px}"
    h = h & ".widget{box-shadow:none;border:1px solid #cbd5e1} @page{size:landscape;margin:10mm}}"
    h = h & "</style></head><body>"
    h = h & "<header><div class='top'><div class='brand'><label class='logo' title='Upload logo'><span id='logoText'>+</span><input id='logoInput' type='file' accept='image/*' hidden></label><h1 id='dashTitle'>" & HtmlEncode(dashTitle) & "</h1></div><select id='currencyCodeSel'><option value=''>No currencyCode</option><option value='BDT'>BDT &#2547;</option><option value='USD'>USD $</option><option value='INR'>INR &#8377;</option><option value='EUR'>EUR &#8364;</option></select><button class='primary' onclick='openLayout()'>Select Layout</button><button onclick='openUpdate()'>Update Data</button><button class='green' onclick='openKPI()"
    h = h & "'>+ KPI</button><button class='green' onclick='openChart()'>+ Chart</button><button class='rose' onclick='openSlicer()'>+ Slicer</button><button onclick='addNote()'>+ Note</button><button class='amber' onclick='openTable()'>+ Table</button><button class='amber' onclick='saveHtml()'>Save</button><button class='danger' onclick='resetDash()'>Reset</button></div></header>"
    h = h & "<main class='wrap'><section class='grid' id='page'></section></main><div class='footer'><b>" & HtmlEncode(dashTitle) & "</b><button onclick='openDataView()'>Data View</button><button onclick='openModeler()'>Data Modeler</button><button class='darkbtn' onclick='print()'>PDF</button></div><div class='modal' id='modal'><div class='box'><div id='modalBody'></div></div></div>"
    h = h & "<script>"
    h = h & "const DATA=" & jsonData & ";"
    h = h & "const START_CUR=" & JsonQuote(currencyCode) & ";"
    h = h & "const SMART_LOCK=" & lockText & ";"
    h = h & "let active=Object.keys(DATA)[0],filters={},pendingSlot=null,dragSlot=null;"
    h = h & "const $=id=>document.getElementById(id);function esc(s){return String(s??'').replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]))}function num(v){return parseFloat(String(v??'').replace(/,/g,''))}"
    h = h & "function cur(){let sel=document.getElementById('currencyCodeSel');if(!sel)return '';let m={BDT:String.fromCharCode(2547),INR:String.fromCharCode(8377),EUR:String.fromCharCode(8364),USD:'$',USDT:'$'};return m[sel.value]||sel.value||''}function fmt(n,c=true){let v=Number(n||0).toLocaleString(undefined,{maximumFractionDigits:2});return c&&cur()?cur()+' '+v:v}"
    h = h & "function hs(){return DATA[active].headers||[]}function raw(){return DATA[active].rows||[]}function nums(){return hs().filter(k=>raw().some(r=>isFinite(num(r[k]))))}function money(){return nums().find(c=>/amount|sales|revenue|price|total|cost|profit|balance|value|paid|due|qty|quantity/i.test(c))||nums()[0]}"
    h = h & "function dateObj(v){let d=new Date(v);if(!isNaN(d))return d;return null}function datePart(v,h){let d=dateObj(v);if(!d||h==='None')return String(v??'');let m=d.getMonth();if(h==='Month')return ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][m];if(h==='Year')return d.getFullYear();return String(v??'')}"
    h = h & "function rows(){let r=raw();Object.keys(filters).forEach(k=>{let f=filters[k];if(f)r=r.filter(x=>String(datePart(x[k],f.h))===String(f.v))});return r}function options(a){return a.map(x=>`<option>${esc(x)}</option>`).join('')}"
    h = h & "function cell(c,v){let n=num(v);let isM=/amount|sales|revenue|price|total|cost|profit|balance|value|paid|due/i.test(c);return isFinite(n)&&isM?fmt(n):esc(v)}function closeModal(){modal.classList.remove('show')}function show(x){modalBody.innerHTML=x;modal.classList.add('show')}"
    h = h & "function slotHTML(t='empty',big=false){return `<div class='slot ${big?'big':''}' data-type='${t}' onclick='slotClick(this)' ondragover='event.preventDefault();this.classList.add(&quot;dragOver&quot;)' ondragleave='this.classList.remove(&quot;dragOver&quot;)' ondrop='dropSlot(this)'>+ Click to Add ${t==='empty'?'Widget':t}</div>`}"
    h = h & "function buildSlots(k=4,c=2,s=2,t=1){page.innerHTML='';for(let i=0;i<s;i++)page.insertAdjacentHTML('beforeend',slotHTML('Slicer'));for(let i=0;i<k;i++)page.insertAdjacentHTML('beforeend',slotHTML('KPI'));for(let i=0;i<c;i++)page.insertAdjacentHTML('beforeend',slotHTML('Chart',true));for(let i=0;i<t;i++)page.insertAdjacentHTML('beforeend',slotHTML('Table',true));saveState()}"
    h = h & "function slotClick(el){if(el.querySelector('.widget'))return;pendingSlot=el;let t=el.dataset.type;if(t==='KPI')openKPI();else if(t==='Chart')openChart();else if(t==='Slicer')openSlicer();else if(t==='Table')openTable();else openLayout()}"
    h = h & "function tools(){return `<div class=tools><button onclick='event.stopPropagation();editWidget(this.closest(&quot;.widget&quot;))'>Edit</button><button onclick='event.stopPropagation();this.closest(&quot;.slot&quot;).innerHTML=&quot;+ Click to Add Widget&quot;;saveState()'>X</button></div>`}"
    h = h & "function put(html){let s=pendingSlot||[...document.querySelectorAll('.slot')].find(x=>!x.querySelector('.widget'));if(!s){page.insertAdjacentHTML('beforeend',slotHTML('empty'));s=page.lastElementChild}s.innerHTML=html;s.querySelector('.widget').draggable=true;pendingSlot=null;wireDrag();refreshAll();saveState()}"
    h = h & "function wireDrag(){document.querySelectorAll('.widget').forEach(w=>{w.ondragstart=e=>{dragSlot=w.closest('.slot')};})}function dropSlot(to){to.classList.remove('dragOver');if(!dragSlot||dragSlot===to)return;let a=dragSlot.innerHTML,b=to.innerHTML;to.innerHTML=a;dragSlot.innerHTML=b;wireDrag();saveState();dragSlot=null}"
    h = h & "function metric(c,op){let r=rows(),a=r.map(x=>num(x[c])).filter(isFinite);if(op==='avg')return a.reduce((x,y)=>x+y,0)/(a.length||1);if(op==='count')return r.length;if(op==='max')return Math.max(...a,0);if(op==='min')return Math.min(...a,0);return a.reduce((x,y)=>x+y,0)}"
    h = h & "function kpiHTML(t,c,op){return `<div class='widget kpi' data-w='kpi' data-title='${esc(t)}' data-col='${esc(c)}' data-op='${op}'>${tools()}<div class=k>${esc(t)}</div><div class=v>${fmt(metric(c,op))}</div><div class=small>${op} of ${esc(c)}</div></div>`}function refreshKPI(w){w.querySelector('.v').textContent=fmt(metric(w.dataset.col,w.dataset.op))}"
    h = h & "function openKPI(w=null){let d=w?.dataset||{},c=d.col||money()||'';show(`<h2>Build KPI</h2><div class=form><div class=field><label>Title</label><input id=kpT value='${esc(d.title||c||'KPI')}'></div><div class=field><label>Column</label><select id=kpC>${options(nums())}"
    h = h & "</select></div><div class=field><label>Aggregation</label><select id=kpO><option value=sum>Sum</option><option value=avg>Average</option><option value=count>Count Rows</option><option value=max>Max</option><option value=min>Min</option></select></div></div><div class=actions><button onclick='closeModal()'>Cancel</button><button class=green onclick='saveKPI()'>Save KPI</button></div>`);kpC.value=c;kpO.value=d.op||'sum';window.editing=w}"
    h = h & "function saveKPI(){let html=kpiHTML(kpT.value,kpC.value,kpO.value);if(window.editing)window.editing.outerHTML=html;else put(html);window.editing=null;closeModal();wireDrag();saveState()}"
    h = h & "function slicerHTML(t,c,h='None'){let vals=[...new Set(raw().map(r=>datePart(r[c],h)).filter(x=>x!==''&&x!=null))].slice(0,30);return `<div class='widget' data-w='slicer' data-title='${esc(t)}' data-col='${esc(c)}' data-h='${h}'>${tools()}<h3 class=title>${esc(t)}</h3><div class=slicerBtns><button class='pill on' onclick='setFilter(&quot;${esc(c)}&quot;,&quot;&quot;,&quot;${h}&quot;)'>All</button>${vals.map(v=>`<button class=pill onclick='setFilter(&quot;${esc(c)}&quot;,&quot;${esc(v)}&quot;,&quot;${h}&quot;)'>${esc(v)}</button>`).join('')}</div></div>`}"
    h = h & "function openSlicer(w=null){let d=w?.dataset||{},c=d.col||hs()[0]||'';show(`<h2>Add Slicer</h2><div class=form><div class=field><label>Title</label><input id=slT value='${esc(d.title||c)}'></div><div class=field><label>Column</label><select id=slC>${options(hs())}</select></div><div class=field><label>Date Group</label><select id=slH><option>None</option><option>Month</option><option>Year</option></select></div></div><div class=actions><button onclick='closeModal()'>Cancel</button><button class=rose onclick='saveSlicer()'>Save Slicer</button></div>`);slC.value=c;slH.value=d.h||'None';window.editing=w}"
    h = h & "function saveSlicer(){let html=slicerHTML(slT.value,slC.value,slH.value);if(window.editing)window.editing.outerHTML=html;else put(html);window.editing=null;closeModal();wireDrag();saveState()}function setFilter(c,v,h){filters[c]=v?{v:v,h:h}:null;refreshAll()}"
    h = h & "function groupData(x,y,agg='sum',dh='None',lim=12){let g={};rows().forEach(r=>{let k=datePart(r[x],dh)||'(blank)',v=num(r[y])||0;if(!g[k])g[k]=[];g[k].push(v)});let a=Object.keys(g).map(k=>{let ar=g[k],v=agg==='avg'?ar.reduce((x,y)=>x+y,0)/ar.length:agg==='count'?ar.length:ar.reduce((x,y)=>x+y,0);return {k,v}});return a.slice(0,lim||a.length)}"
    h = h & "function chartHTML(t,type,x,y,agg,dh,xl,yl,lim){return `<div class='widget' data-w='chart' data-title='${esc(t)}' data-type='${type}' data-x='${esc(x)}' data-y='${esc(y)}' data-agg='${agg}' data-dh='${dh}' data-xl='${esc(xl)}' data-yl='${esc(yl)}' data-lim='${lim}'>${tools()}<h3 class=title>${esc(t)}</h3><div class=chartBox></div></div>`}"
    h = h & "function drawChart(w){let a=groupData(w.dataset.x,w.dataset.y,w.dataset.agg,w.dataset.dh,+w.dataset.lim||12),type=w.dataset.type,W=760,H=290,p=42,max=Math.max(...a.map(o=>o.v),1),base=H-p;if(!a.length){w.querySelector('.chartBox').innerHTML='No data';return}let step=(W-p*2)/a.length,bw=step*.62;"
    h = h & "let bars=a.map((o,i)=>{let bh=o.v/max*(H-p*2),x=p+i*step+step*.19,y=base-bh;return `<rect class=bar x='${x}' y='${y}' width='${bw}' height='${bh}' rx='5'></rect><text class=label x='${x+bw/2}' y='${y-6}' text-anchor='middle'>${fmt(o.v)}</text>`}).join('');let pts=a.map((o,i)=>[p+i*step+step/2,base-o.v/max*(H-p*2)]);"
    h = h & "let line=`<polyline class=line points='${pts.map(p=>p.join(',')).join(' ')}'/>`+pts.map((p,i)=>`<circle cx='${p[0]}' cy='${p[1]}' r='4' fill='var(--acc)'/><text class=label x='${p[0]}' y='${p[1]-8}' text-anchor='middle'>${fmt(a[i].v)}</text>`).join('');let labs=a.map((o,i)=>`<text class=label x='${p+i*step+step/2}' y='${H-10}' text-anchor='middle'>${esc(o.k).slice(0,10)}</text>`).join('');"
    h = h & "w.querySelector('.chartBox').innerHTML=`<svg viewBox='0 0 ${W} ${H}'><text class=label x='8' y='18'>${esc(w.dataset.yl||w.dataset.y)}</text><line class=axis x1='${p}' y1='${base}' x2='${W-p}' y2='${base}'/>${type==='line'||type==='area'?line:bars}${labs}<text class=label x='${W/2}' y='${H-1}' text-anchor='middle'>${esc(w.dataset.xl||w.dataset.x)}</text></svg>`}"
    h = h & "function openChart(w=null){let d=w?.dataset||{},x=d.x||hs()[0]||'',y=d.y||money()||'';show(`<h2>Build Analytics Chart</h2><div class=form><div class=field><label>Title</label><input id=chT value='${esc(d.title||'Month Wise Sales')}'></div><div class=field><label>Chart Type</label><select id=chType><option value=bar>Column</option><option value=line>Line</option><option value=area>Area</option></select></div><div class=field><label>Aggregation</label><select id=chAgg><option value=sum>Sum Total</option><option value=avg>Average</option><option value=count>Count</option></select></div><div class=field><label>"
    h = h & "X Axis</label><select id=chX>${options(hs())}</select></div><div class=field><label>Y Axis</label><select id=chY>${options(nums())}</select></div><div class=field><label>Date Hierarchy</label><select id=chDh><option>None</option><option>Month</option><option>Year</option></select></div><div class=field><label>X Label</label><input id=chXL value='${esc(d.xl||x)}'></div><div class=field><label>Y Label</label><input id=chYL value='${esc(d.yl||y)}'></div><div class=field><label>Top N</label><input id=chLim value='${d.lim||12}"
    h = h & "'></div></div><div class=actions><button onclick='closeModal()'>Cancel</button><button class=green onclick='saveChart()'>Generate Chart</button></div>`);chX.value=x;chY.value=y;chType.value=d.type||'bar';chAgg.value=d.agg||'sum';chDh.value=d.dh||'None';window.editing=w}"
    h = h & "function saveChart(){let html=chartHTML(chT.value,chType.value,chX.value,chY.value,chAgg.value,chDh.value,chXL.value,chYL.value,chLim.value);if(window.editing)window.editing.outerHTML=html;else put(html);window.editing=null;closeModal();wireDrag();refreshAll();saveState()}"
    h = h & "function tableHTML(t,row,col,val,agg){return `<div class='widget' data-w='table' data-title='${esc(t)}' data-row='${esc(row)}' data-col='${esc(col)}' data-val='${esc(val)}' data-agg='${agg}'>${tools()}<h3 class=title>${esc(t)}</h3><div class=tableWrap><table></table></div></div>`}"
    h = h & "function drawTable(w){let r=w.dataset.row,c=w.dataset.col,v=w.dataset.val,agg=w.dataset.agg,rs=rows(),cols=c?[...new Set(rs.map(x=>x[c]))].slice(0,20):[],map={};rs.forEach(x=>{let rk=x[r]||'(blank)',ck=c?(x[c]||'(blank)'):'Total';if(!map[rk])map[rk]={};if(!map[rk][ck])map[rk][ck]=[];map[rk][ck].push(num(x[v])||0)}"
    h = h & ");let head='<tr><th>'+esc(r)+'</th>'+cols.map(x=>'<th>'+esc(x)+'</th>').join('')+'<th>Grand Total</th></tr>';let body=Object.keys(map).slice(0,60).map(rk=>{let total=0,td=cols.map(ck=>{let ar=map[rk][ck]||[],n=agg==='avg'?ar.reduce((a,b)=>a+b,0)/(ar.length||1):ar.reduce((a,b)=>a+b,0);total+=n;return '<td>'+fmt(n)+'</td>'}).join('');return '<tr><th>'+esc(rk)+'</th>'+td+'<td>'+fmt(total)+'</td></tr>'}).join('');w.querySelector('table').innerHTML='<thead>'+head+'</thead><tbody>'+body+'</tbody>'}"
    h = h & "function openTable(w=null){let d=w?.dataset||{},r=d.row||hs()[0]||'',v=d.val||money()||'';show(`<h2>Smart Table / Pivot</h2><div class=form><div class=field><label>Title</label><input id=tbT value='${esc(d.title||'Summary Table')}'></div><div class=field><label>Rows</label><select id=tbR>${options(hs())}</select></div><div class=field><label>Columns</label><select id=tbC><option value=''>None</option>${options(hs())}</select></div><div class=field><label>Value</label><select id=tbV>${options(nums())}"
    h = h & "</select></div><div class=field><label>Calculation</label><select id=tbA><option value=sum>Sum</option><option value=avg>Average</option></select></div></div><div class=actions><button onclick='closeModal()'>Cancel</button><button class=amber onclick='saveTable()'>Generate Table</button></div>`);tbR.value=r;tbC.value=d.col||'';tbV.value=v;tbA.value=d.agg||'sum';window.editing=w}"
    h = h & "function saveTable(){let html=tableHTML(tbT.value,tbR.value,tbC.value,tbV.value,tbA.value);if(window.editing)window.editing.outerHTML=html;else put(html);window.editing=null;closeModal();wireDrag();refreshAll();saveState()}"
    h = h & "function addNote(){put(`<div class='widget note' data-w='note'>${tools()}<input class=noteTitle value='Management Note'><textarea>Write insight here...</textarea></div>`)}function editWidget(w){let t=w.dataset.w;if(t==='kpi')openKPI(w);if(t==='slicer')openSlicer(w);if(t==='chart')openChart(w);if(t==='table')openTable(w)}"
    h = h & "function refreshAll(){document.querySelectorAll('[data-w=kpi]').forEach(refreshKPI);document.querySelectorAll('[data-w=chart]').forEach(drawChart);document.querySelectorAll('[data-w=table]').forEach(drawTable)}function saveState(){localStorage.setItem('pd:'+document.title,page.innerHTML)}function loadState(){let s=localStorage.getItem('pd:'+document.title);if(s){page.innerHTML=s;wireDrag();refreshAll();return true}return false}"
    h = h & "function openLayout(){show(`<h2>Professional Layout Builder</h2><div class=section><h3>Custom Grid Builder</h3><div class=form><div class=field><label>KPI Cards</label><input id=lk value=4></div><div class=field><label>Charts</label><input id=lc value=2></div><div class=field><label>Slicers</label><input id=ls value=4></div><div class=field><label>Tables</label><input id=lt value=1></div></div><div class=actions><button class=primary onclick='buildSlots(+lk.value,+lc.value,+ls.value,+lt.value);"
    h = h & "closeModal()'>Build Grid</button></div></div><div class=section><h3>Quick Pro Templates</h3><div class=templates>${['Balanced Classic','Hero Chart','Sidebar Focus','Split Panels','Executive Grid','Analytics Pro'].map((n,i)=>`<div class=tpl onclick='buildSlots(${i===5?6:4},${i===1?1:2},${i===2?1:4},1);closeModal()'><div class=sk><i></i><i></i><i></i><i></i><i style='grid-column:span 2'></i><i style='grid-column:span 2'></i></div><b>${n}</b><p class=small>Click to apply</p></div>`).join('')}</div></div><div class=actions><button onclick='closeModal()'>Close</button></div>`)}"
    h = h & "function openUpdate(){show(`<h2>Update Data</h2><p>This HTML already contains all Excel data. To refresh data, generate dashboard again from Excel.</p><div class=actions><button onclick='closeModal()'>Close</button></div>`)}function openDataView(){let hs2=hs(),rs=raw().slice(0,50);show(`<h2>Data View</h2><p class=small>Showing first 50 rows. Full data stays embedded in backend JS data.</p><div class=tableWrap><table><thead><tr>${hs2.map(h=>`<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rs.map(r=>`<tr>${hs2.map(h=>`<td>${cell(h,r[h])}</td>`).join('')}</tr>`).join('')}</tbody></table></div><div class=actions><button onclick='closeModal()'>Close</button></div>`)}"
    h = h & "function openModeler(){show(`<h2>Data Modeler</h2><p>No relation feature is enabled. This dashboard uses the embedded Excel data tables only.</p><div class=actions><button onclick='closeModal()'>Close</button></div>`)}function saveHtml(){if(SMART_LOCK){alert('Save disabled by Smart Lock');return}let html='<!doctype html>'+document.documentElement.outerHTML;if(window.showSaveFilePicker){showSaveFilePicker({suggestedName:document.title+'.html',types:[{description:'HTML',accept:{'text/html':['.html']}}]}).then(async h=>{let w=await h.createWritable();await w.write(html);await w.close()})}"
    h = h & "else{let a=document.createElement('a');a.href=URL.createObjectURL(new Blob([html],{type:'text/html'}));a.download=document.title+'.html';a.click()}}"
    h = h & "function resetDash(){if(confirm('Remove all widgets and return blank?')){localStorage.removeItem('pd:'+document.title);page.innerHTML='';buildSlots(0,0,0,0)}}function normCur(x){x=String(x||'').trim();let sym={'\u09F3':'BDT','\u20B9':'INR','\u20AC':'EUR','$':'USD'};return sym[x]||x}"
    h = h & "currencyCodeSel.value=normCur(START_CUR);logoInput.onchange=e=>{let f=e.target.files[0];if(!f)return;let r=new FileReader();r.onload=x=>{logoText.innerHTML='<img src='+x.target.result+'>';};r.readAsDataURL(f)};"
    h = h & "if(!loadState())buildSlots(4,2,4,1);currencyCodeSel.onchange=()=>{refreshAll();saveState()};"
    h = h & "</script></body></html>"

    BuildDashboardHTML = h
End Function

Private Function HtmlEncode(ByVal text As String) As String
    text = Replace(text, "&", "&amp;")
    text = Replace(text, "<", "&lt;")
    text = Replace(text, ">", "&gt;")
    text = Replace(text, """", "&quot;")
    HtmlEncode = text
End Function
-------------------------------------------------------------------------------

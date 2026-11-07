Attribute VB_Name = "modDashTemplate"
Option Explicit

'===============================================================================
'  POWER DASHBOARD - HTML TEMPLATE  (SINGLE FILE MODULE)
'-------------------------------------------------------------------------------
'  Ei module e dashboard er puro HTML/JavaScript template thake.
'  GetDashboardTemplate() ke call korle sob niche-er Sec01..Sec20 ekshathe
'  joro hoye ekta puro HTML return kore; modPowerDashboardEngine seita diye
'  rawData + aiConfig boshaiye .html file toiri kore.
'-------------------------------------------------------------------------------
'  SECTION INDEX (kono-er sathe ki ache):
'    Sec01_HeadCss            : Document head, CDN libraries (ECharts/Three.js/GridStack/XLSX) and the complete CSS theme [40477 chars, 3 chunks]
'    Sec02_BasicModals        : Add Slicer, Update Data (Excel) and Label Settings modal markup [8458 chars]
'    Sec03_AddKpiModal        : Add KPI modal (title, column, aggregation, icon/image picker) [15865 chars]
'    Sec04_EditKpiModal       : Edit KPI modal [15898 chars]
'    Sec05_CreateChartModal   : Create Chart modal (type, X/Y column, hierarchy, labels) [5527 chars]
'    Sec06_EditChartModal     : Edit Chart modal [5511 chars]
'    Sec07_LayoutModal        : Layout picker modal - 6 preset layouts + custom KPI/chart/slicer layout [5690 chars]
'    Sec08_HeaderLogo         : Dashboard header, toolbar buttons (Layout/Update/KPI/Chart/Slicer/Save/Reset/Deploy) and the embedded company logo bitmap [281907 chars, 18 chunks]
'    Sec09_GridArea           : 3D background canvas, slicer bar and the GridStack container (#mainGrid) [27025 chars, 2 chunks]
'    Sec10_Footer             : Footer bar with the live theme swatch switcher [1833 chars]
'    Sec11_ScriptCore         : MAIN SCRIPT: RAWDATA/AICONFIG injection markers, 3D particle background, icon library, GridStack init, auto-save [12484 chars]
'    Sec12_LayoutEngine       : Layout engine: Edit-Layout toggle, placeholders, applyLayout / applyCustomLayout [5924 chars]
'    Sec13_DataEngine         : Update data from Excel, date detection, global filters, logo, aggregation and number formatting helpers [14893 chars]
'    Sec14_Renderers          : renderKPIs + renderCharts (ECharts options: bar/line/area/pie/doughnut, toolbox, labels) [13034 chars]
'    Sec15_ChartDialogs       : Delete KPI/chart (Smart Lock guard) and Create/Edit chart dialogs [6371 chars]
'    Sec16_Slicers            : Slicer add / render / delete and the row filter engine [46251 chars, 3 chunks]
'    Sec17_SaveHtml           : Auto-save timer and Save HTML export (File System Access API + download fallback) [5356 chars]
'    Sec18_UpdateAndLogo      : updateAll, resetFilters, closeModals, label settings, logo upload [6691 chars]
'    Sec19_Themes             : Theme definitions: Light, Dark, Ocean, Midnight, Animated, Forest, Rose, Navy [10396 chars]
'    Sec20_GithubDeploy       : GitHub Pages deploy modal + deployToGithub(), document close [8456 chars]
'-------------------------------------------------------------------------------
'  NOTE: Prota section ekta alada Function. Kintu module ekta-i (one file).
'  Chorom chunk gulo kete 18000 char-er paket kora ache, jate VBA-er
'  64KB per-procedure limit na buge.
'===============================================================================

Public Function GetDashboardTemplate() As String
    Dim s As String
    s = s & Sec01_HeadCss_1()
    s = s & Sec01_HeadCss_2()
    s = s & Sec01_HeadCss_3()
    s = s & Sec02_BasicModals()
    s = s & Sec03_AddKpiModal()
    s = s & Sec04_EditKpiModal()
    s = s & Sec05_CreateChartModal()
    s = s & Sec06_EditChartModal()
    s = s & Sec07_LayoutModal()
    s = s & Sec08_HeaderLogo_1()
    s = s & Sec08_HeaderLogo_2()
    s = s & Sec08_HeaderLogo_3()
    s = s & Sec08_HeaderLogo_4()
    s = s & Sec08_HeaderLogo_5()
    s = s & Sec08_HeaderLogo_6()
    s = s & Sec08_HeaderLogo_7()
    s = s & Sec08_HeaderLogo_8()
    s = s & Sec08_HeaderLogo_9()
    s = s & Sec08_HeaderLogo_10()
    s = s & Sec08_HeaderLogo_11()
    s = s & Sec08_HeaderLogo_12()
    s = s & Sec08_HeaderLogo_13()
    s = s & Sec08_HeaderLogo_14()
    s = s & Sec08_HeaderLogo_15()
    s = s & Sec08_HeaderLogo_16()
    s = s & Sec08_HeaderLogo_17()
    s = s & Sec08_HeaderLogo_18()
    s = s & Sec09_GridArea_1()
    s = s & Sec09_GridArea_2()
    s = s & Sec10_Footer()
    s = s & Sec11_ScriptCore()
    s = s & Sec12_LayoutEngine()
    s = s & Sec13_DataEngine()
    s = s & Sec14_Renderers()
    s = s & Sec15_ChartDialogs()
    s = s & Sec16_Slicers_1()
    s = s & Sec16_Slicers_2()
    s = s & Sec16_Slicers_3()
    s = s & Sec17_SaveHtml()
    s = s & Sec18_UpdateAndLogo()
    s = s & Sec19_Themes()
    s = s & Sec20_GithubDeploy()
    GetDashboardTemplate = s
End Function

'--- Sec01_HeadCss_1 : Document head, CDN libraries (ECharts/Three.js/GridStack/XLSX) and the complete CSS theme (chunk 1 of 3)
Private Function Sec01_HeadCss_1() As String
    Dim s As String
    s = s & "<!DOCTYPE html>" & vbLf
    s = s & "<html lang=""en""" & vbLf
    s = s & "  style=""--body-bg: #081b29; --card-bg: rgba(8, 27, 41, 0.75); --text-col: #ffffff; --border-col: rgba(0, 238, 255, 0.3); --grid-col: rgba(255, 255, 255, 0.05); --axis-col: #ededed; --shadow-col: rgba(0, 238, 255, 0.2); --header-bg: rgba(8, 27, 41, 0.85); --input-bg: #112e42; --scroll-thumb: #0ef; --kpi-title-col: #ffffff; --kpi-val-col: #0ef; --chart-header-col: #ffffff;"">" & vbLf
    s = s & "" & vbLf
    s = s & "<head>" & vbLf
    s = s & "  <meta charset=""UTF-8"">" & vbLf
    s = s & "  <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">" & vbLf
    s = s & "  <title>XLBooster Dashboard</title>" & vbLf
    s = s & "  <link rel=""stylesheet"" href=""https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css"">" & vbLf
    s = s & "  <script src=""https://cdn.jsdelivr.net/npm/echarts@5.5.0/dist/echarts.min.js""></script>" & vbLf
    s = s & "  <script src=""https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js""></script>" & vbLf
    s = s & "  <script src=""https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js""></script>" & vbLf
    s = s & "  <script src=""https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js""></script>" & vbLf
    s = s & "  <link href=""https://cdn.jsdelivr.net/npm/gridstack@10.1.1/dist/gridstack.min.css"" rel=""stylesheet"">" & vbLf
    s = s & "  <script src=""https://cdn.jsdelivr.net/npm/gridstack@10.1.1/dist/gridstack-all.js""></script>" & vbLf
    s = s & "  <link href=""https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&amp;display=swap"" rel=""stylesheet"">" & vbLf
    s = s & "  <link href=""https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&amp;display=swap""" & vbLf
    s = s & "    rel=""stylesheet"">" & vbLf
    s = s & "  <style>" & vbLf
    s = s & "    :root {" & vbLf
    s = s & "      --body-bg: #f8fafc;" & vbLf
    s = s & "      --card-bg: #ffffff;" & vbLf
    s = s & "      --text-col: #0f172a;" & vbLf
    s = s & "      --border-col: #e2e8f0;" & vbLf
    s = s & "      --grid-col: #f1f5f9;" & vbLf
    s = s & "      --axis-col: #64748b;" & vbLf
    s = s & "      --shadow-col: rgba(15, 23, 42, 0.06);" & vbLf
    s = s & "      --header-bg: rgba(255, 255, 255, 0.85);" & vbLf
    s = s & "      --input-bg: #f1f5f9;" & vbLf
    s = s & "      --scroll-thumb: #cbd5e1;" & vbLf
    s = s & "      --kpi-title-col: #64748b;" & vbLf
    s = s & "      --kpi-val-col: #4f46e5;" & vbLf
    s = s & "      --chart-header-col: #1e293b;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    @keyframes gradientBG {" & vbLf
    s = s & "      0% {" & vbLf
    s = s & "        background-position: 0% 50%;" & vbLf
    s = s & "      }" & vbLf
    s = s & "" & vbLf
    s = s & "      50% {" & vbLf
    s = s & "        background-position: 100% 50%;" & vbLf
    s = s & "      }" & vbLf
    s = s & "" & vbLf
    s = s & "      100% {" & vbLf
    s = s & "        background-position: 0% 50%;" & vbLf
    s = s & "      }" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    @keyframes floatCard {" & vbLf
    s = s & "" & vbLf
    s = s & "      0%," & vbLf
    s = s & "      100% {" & vbLf
    s = s & "        transform: translateY(0px);" & vbLf
    s = s & "      }" & vbLf
    s = s & "" & vbLf
    s = s & "      50% {" & vbLf
    s = s & "        transform: translateY(-8px);" & vbLf
    s = s & "      }" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    body.animated-theme {" & vbLf
    s = s & "      background-color: var(--body-bg) !important;" & vbLf
    s = s & "      font-family: 'Poppins', sans-serif !important;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    #canvas-container {" & vbLf
    s = s & "      position: fixed;" & vbLf
    s = s & "      top: 0;" & vbLf
    s = s & "      left: 0;" & vbLf
    s = s & "      width: 100vw;" & vbLf
    s = s & "      height: 100vh;" & vbLf
    s = s & "      z-index: -1;" & vbLf
    s = s & "      pointer-events: none;" & vbLf
    s = s & "      opacity: 0;" & vbLf
    s = s & "      transition: opacity 1s ease-in-out;" & vbLf
    s = s & "      display: none;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    body.animated-theme #canvas-container {" & vbLf
    s = s & "      display: block;" & vbLf
    s = s & "      opacity: 1;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    body.animated-theme .kpi-card," & vbLf
    s = s & "    body.animated-theme .chart-card {" & vbLf
    s = s & "      background: var(--card-bg) !important;" & vbLf
    s = s & "      backdrop-filter: blur(16px) !important;" & vbLf
    s = s & "      border: 1px solid rgba(0, 238, 255, 0.1) !important;" & vbLf
    s = s & "      transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    body.animated-theme .kpi-card:hover," & vbLf
    s = s & "    body.animated-theme .chart-card:hover {" & vbLf
    s = s & "      border-color: #0ef !important;" & vbLf
    s = s & "      box-shadow: 0 0 20px rgba(0, 238, 255, 0.4) !important;" & vbLf
    s = s & "      animation: floatCard 3s ease-in-out infinite !important;" & vbLf
    s = s & "      z-index: 10;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    * {" & vbLf
    s = s & "      box-sizing: border-box;" & vbLf
    s = s & "      outline: none;" & vbLf
    s = s & "      margin: 0;" & vbLf
    s = s & "      padding: 0;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    body {" & vbLf
    s = s & "      background-color: var(--body-bg);" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "      font-family: 'Inter', sans-serif;" & vbLf
    s = s & "      overflow-x: hidden;" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      flex-direction: column;" & vbLf
    s = s & "      min-height: 100vh;" & vbLf
    s = s & "      padding-bottom: 64px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    ::-webkit-scrollbar {" & vbLf
    s = s & "      width: 8px;" & vbLf
    s = s & "      height: 8px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    ::-webkit-scrollbar-track {" & vbLf
    s = s & "      background: var(--body-bg);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    ::-webkit-scrollbar-thumb {" & vbLf
    s = s & "      background: var(--scroll-thumb);" & vbLf
    s = s & "      border-radius: 4px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dash-header {" & vbLf
    s = s & "      position: sticky;" & vbLf
    s = s & "      top: 0;" & vbLf
    s = s & "      z-index: 1000;" & vbLf
    s = s & "      background: var(--header-bg);" & vbLf
    s = s & "      backdrop-filter: blur(12px);" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      flex-direction: column;" & vbLf
    s = s & "      border-bottom: 1px solid var(--border-col);" & vbLf
    s = s & "      box-shadow: 0 4px 20px var(--shadow-col);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .header-top {" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      justify-content: space-between;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "      padding: 15px 30px;" & vbLf
    s = s & "      width: 100%;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dash-title {" & vbLf
    s = s & "      font-size: 22px;" & vbLf
    s = s & "      font-weight: 700;" & vbLf
    s = s & "      letter-spacing: -0.5px;" & vbLf
    s = s & "      text-transform: uppercase;" & vbLf
    s = s & "      background: linear-gradient(135deg, #6366f1 0%, #a855f7 100%);" & vbLf
    s = s & "      -webkit-background-clip: text;" & vbLf
    s = s & "      -webkit-text-fill-color: transparent;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .header-actions {" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      gap: 10px;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .btn-save {" & vbLf
    s = s & "      background: linear-gradient(135deg, #4f46e5 0%, #4338ca 100%);" & vbLf
    s = s & "      color: white;" & vbLf
    s = s & "      border: none;" & vbLf
    s = s & "      padding: 10px 16px;" & vbLf
    s = s & "      border-radius: 8px;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      font-weight: 600;" & vbLf
    s = s & "      font-size: 12px;" & vbLf
    s = s & "      box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);" & vbLf
    s = s & "      transition: all 0.2s ease;" & vbLf
    s = s & "      white-space: nowrap;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .btn-save:hover {" & vbLf
    s = s & "      transform: translateY(-1px);" & vbLf
    s = s & "      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .btn-create {" & vbLf
    s = s & "      background: linear-gradient(135deg, #10b981 0%, #059669 100%);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .btn-export {" & vbLf
    s = s & "      background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .btn-update {" & vbLf
    s = s & "      background: linear-gradient(135deg, #0ea5e9 0%, #2563eb 100%);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .btn-slicer {" & vbLf
    s = s & "      background: linear-gradient(135deg, #f43f5e 0%, #e11d48 100%);" & vbLf
    s = s & "      box-shadow: 0 2px 10px rgba(225, 29, 72, 0.3);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .btn-table {" & vbLf
    s = s & "      background: linear-gradient(135deg, #06b6d4 0%, #0e7490 100%);" & vbLf
    s = s & "      box-shadow: 0 2px 10px rgba(6, 182, 212, 0.3);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    /* ---------- Add Table : data table widget ---------- */" & vbLf
    s = s & "    .table-card {" & vbLf
    s = s & "      position: relative;" & vbLf
    s = s & "      width: 100%;" & vbLf
    s = s & "      height: 100%;" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      flex-direction: column;" & vbLf
    s = s & "      background-color: var(--card-bg);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      border-radius: 16px;" & vbLf
    s = s & "      box-shadow: 0 4px 6px -1px var(--shadow-col);" & vbLf
    s = s & "      overflow: hidden;" & vbLf
    s = s & "      transition: all 0.3s ease;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .table-card:hover {" & vbLf
    s = s & "      border-color: var(--kpi-val-col);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .table-head {" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "      justify-content: space-between;" & vbLf
    s = s & "      gap: 8px;" & vbLf
    s = s & "      padding: 12px 44px 12px 16px;" & vbLf
    s = s & "      background: linear-gradient(135deg, rgba(6, 182, 212, 0.14), rgba(14, 116, 144, 0.06));" & vbLf
    s = s & "      border-bottom: 1px solid var(--border-col);" & vbLf
    s = s & "      flex-shrink: 0;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .table-title {" & vbLf
    s = s & "      font-size: 13px;" & vbLf
    s = s & "      font-weight: 700;" & vbLf
    s = s & "      letter-spacing: 0.4px;" & vbLf
    s = s & "      text-transform: uppercase;" & vbLf
    s = s & "      color: var(--chart-header-col);" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      white-space: nowrap;" & vbLf
    s = s & "      overflow: hidden;" & vbLf
    s = s & "      text-overflow: ellipsis;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .table-tools {" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "      gap: 6px;" & vbLf
    s = s & "      flex-shrink: 0;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .tbl-search," & vbLf
    s = s & "    .tbl-finput {" & vbLf
    s = s & "      background: var(--input-bg);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "      border-radius: 6px;" & vbLf
    s = s & "      padding: 5px 8px;" & vbLf
    s = s & "      font-size: 11px;" & vbLf
    s = s & "      outline: none;" & vbLf
    s = s & "      min-width: 0;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .tbl-search {" & vbLf
    s = s & "      width: 120px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .tbl-sdate {" & vbLf
    s = s & "      width: 185px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .tbl-search:focus," & vbLf
    s = s & "    .tbl-finput:focus {" & vbLf
    s = s & "      border-color: #06b6d4;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .tbl-btn {" & vbLf
    s = s & "      border: none;" & vbLf
    s = s & "      border-radius: 6px;" & vbLf
    s = s & "      width: 26px;" & vbLf
    s = s & "      height: 26px;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      font-size: 12px;" & vbLf
    s = s & "      font-weight: 700;" & vbLf
    s = s & "      color: #fff;" & vbLf
    s = s & "      line-height: 1;" & vbLf
    s = s & "      transition: transform 0.15s;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .tbl-btn:hover {" & vbLf
    s = s & "      transform: translateY(-1px);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .tbl-btn.csv { background: #10b981; }" & vbLf
    s = s & "    .tbl-btn.pdf { background: #ef4444; }" & vbLf
    s = s & "    .tbl-btn.edit { background: #8b5cf6; }" & vbLf
    s = s & "" & vbLf
    s = s & "    .delete-table {" & vbLf
    s = s & "      position: absolute;" & vbLf
    s = s & "      top: 10px;" & vbLf
    s = s & "      right: 10px;" & vbLf
    s = s & "      background: transparent;" & vbLf
    s = s & "      border: none;" & vbLf
    s = s & "      color: #ef4444;" & vbLf
    s = s & "      font-size: 18px;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      opacity: 0;" & vbLf
    s = s & "      transition: opacity 0.2s;" & vbLf
    s = s & "      z-index: 10;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .table-card:hover .delete-table {" & vbLf
    s = s & "      opacity: 1;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .table-scroll {" & vbLf
    s = s & "      flex: 1;" & vbLf
    s = s & "      overflow: auto;" & vbLf
    s = s & "      min-height: 0;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .table-scroll::-webkit-scrollbar," & vbLf
    s = s & "    .table-body::-webkit-scrollbar {" & vbLf
    s = s & "      width: 6px;" & vbLf
    s = s & "      height: 6px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .table-scroll::-webkit-scrollbar-thumb," & vbLf
    s = s & "    .table-body::-webkit-scrollbar-thumb {" & vbLf
    s = s & "      background: var(--scroll-thumb);" & vbLf
    s = s & "      border-radius: 4px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    table.dash-table {" & vbLf
    s = s & "      width: 100%;" & vbLf
    s = s & "      border-collapse: collapse;" & vbLf
    s = s & "      font-size: 12px;" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dash-table th," & vbLf
    s = s & "    .dash-table td {" & vbLf
    s = s & "      padding: 7px 10px;" & vbLf
    s = s & "      text-align: left;" & vbLf
    s = s & "      border-bottom: 1px solid var(--border-col);" & vbLf
    s = s & "      white-space: nowrap;" & vbLf
    s = s & "      max-width: 220px;" & vbLf
    s = s & "      overflow: hidden;" & vbLf
    s = s & "      text-overflow: ellipsis;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dash-table th {" & vbLf
    s = s & "      position: sticky;" & vbLf
    s = s & "      top: 0;" & vbLf
    s = s & "      z-index: 2;" & vbLf
    s = s & "      background: var(--card-bg);" & vbLf
    s = s & "      font-weight: 700;" & vbLf
    s = s & "      font-size: 11px;" & vbLf
    s = s & "      letter-spacing: 0.4px;" & vbLf
    s = s & "      text-transform: uppercase;" & vbLf
    s = s & "      color: var(--chart-header-col);" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      user-select: none;" & vbLf
    s = s & "      border-bottom: 2px solid var(--kpi-val-col);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dash-table th .sort-arr { font-size: 9px; opacity: 0.55; margin-left: 3px; }" & vbLf
    s = s & "    .dash-table th.sorted .sort-arr { opacity: 1; color: var(--kpi-val-col); }" & vbLf
    s = s & "" & vbLf
    s = s & "    .sort-btns {" & vbLf
    s = s & "      display: inline-flex;" & vbLf
    s = s & "      gap: 3px;" & vbLf
    s = s & "      margin-left: 5px;" & vbLf
    s = s & "      vertical-align: middle;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .sort-btns a {" & vbLf
    s = s & "      font-size: 8.5px;" & vbLf
    s = s & "      font-weight: 800;" & vbLf
    s = s & "      letter-spacing: 0;" & vbLf
    s = s & "      text-transform: none;" & vbLf
    s = s & "      color: var(--axis-col);" & vbLf
    s = s & "      background: var(--input-bg);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      border-radius: 4px;" & vbLf
    s = s & "      padding: 1px 4px;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      text-decoration: none;" & vbLf
    s = s & "      opacity: 0.6;" & vbLf
    s = s & "      white-space: nowrap;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .sort-btns a:hover { opacity: 1; border-color: #06b6d4; color: #06b6d4; }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dash-table tbody tr:hover { background: rgba(127, 127, 127, 0.08); }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dash-table td.num," & vbLf
    s = s & "    .dash-table th.num { text-align: right; }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dash-table tfoot td {" & vbLf
    s = s & "      font-weight: 700;" & vbLf
    s = s & "      background: rgba(127, 127, 127, 0.1);" & vbLf
    s = s & "      border-top: 2px solid var(--kpi-val-col);" & vbLf
    s = s & "      position: sticky;" & vbLf
    s = s & "      bottom: 0;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .table-filterrow td {" & vbLf
    s = s & "      padding: 4px 6px;" & vbLf
    s = s & "      background: var(--input-bg);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .table-scroll .dash-table th:nth-child(1)," & vbLf
    s = s & "    .table-scroll .dash-table tbody td:nth-child(1)," & vbLf
    s = s & "    .table-scroll .dash-table tfoot td:nth-child(1)," & vbLf
    s = s & "    .table-scroll .table-filterrow td:nth-child(1) {" & vbLf
    s = s & "      position: sticky;" & vbLf
    s = s & "      left: 0;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .table-scroll .dash-table th:nth-child(2)," & vbLf
    s = s & "    .table-scroll .dash-table tbody td:nth-child(2)," & vbLf
    s = s & "    .table-scroll .dash-table tfoot td:nth-child(2)," & vbLf
    s = s & "    .table-scroll .table-filterrow td:nth-child(2) {" & vbLf
    s = s & "      position: sticky;" & vbLf
    s = s & "      left: var(--fz1, 96px);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .table-scroll .dash-table tbody td:nth-child(1)," & vbLf
    s = s & "    .table-scroll .dash-table tbody td:nth-child(2) {" & vbLf
    s = s & "      z-index: 3;" & vbLf
    s = s & "      background: var(--card-bg);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .table-scroll .dash-table th:nth-child(1)," & vbLf
    s = s & "    .table-scroll .dash-table th:nth-child(2) {" & vbLf
    s = s & "      z-index: 4;" & vbLf
    s = s & "      background: var(--card-bg);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .table-scroll .table-filterrow td:nth-child(1)," & vbLf
    s = s & "    .table-scroll .table-filterrow td:nth-child(2) {" & vbLf
    s = s & "      z-index: 3;" & vbLf
    s = s & "      background: var(--input-bg);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .table-scroll .dash-table tfoot td:nth-child(1)," & vbLf
    s = s & "    .table-scroll .dash-table tfoot td:nth-child(2) {" & vbLf
    s = s & "      z-index: 3;" & vbLf
    s = s & "      background: color-mix(in srgb, var(--card-bg) 90%, #808080 10%);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .table-scroll .dash-table tbody tr:hover td:nth-child(1)," & vbLf
    s = s & "    .table-scroll .dash-table tbody tr:hover td:nth-child(2) {" & vbLf
    s = s & "      background: color-mix(in srgb, var(--card-bg) 92%, #808080 8%);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .table-filterrow .tbl-finput { width: 100%; box-sizing: border-box; }" & vbLf
    s = s & "" & vbLf
    s = s & "    .table-foot {" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "      justify-content: space-between;" & vbLf
    s = s & "      gap: 8px;" & vbLf
    s = s & "      padding: 7px 12px;" & vbLf
    s = s & "      font-size: 11px;" & vbLf
    s = s & "      color: var(--axis-col);" & vbLf
    s = s & "      border-top: 1px solid var(--border-col);" & vbLf
    s = s & "      background: var(--card-bg);" & vbLf
    s = s & "      flex-shrink: 0;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .page-btn {" & vbLf
    s = s & "      background: var(--input-bg);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "      border-radius: 6px;" & vbLf
    s = s & "      padding: 3px 9px;" & vbLf
    s = s & "      font-size: 11px;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .page-btn:disabled { opacity: 0.4; cursor: default; }" & vbLf
    s = s & "    .page-btn:not(:disabled):hover { border-color: #06b6d4; }" & vbLf
    s = s & "" & vbLf
    s = s & "    .tbl-empty {" & vbLf
    s = s & "      padding: 26px 12px;" & vbLf
    s = s & "      text-align: center;" & vbLf
    s = s & "      color: var(--axis-col);" & vbLf
    s = s & "      font-size: 12px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .col-picker {" & vbLf
    s = s & "      display: grid;" & vbLf
    s = s & "      grid-template-columns: repeat(auto-fill, minmax(140px, 1fr));" & vbLf
    s = s & "      gap: 6px;" & vbLf
    s = s & "      max-height: 150px;" & vbLf
    s = s & "      overflow-y: auto;" & vbLf
    s = s & "      padding: 8px;" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      border-radius: 8px;" & vbLf
    s = s & "      background: var(--input-bg);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .col-pick-item {" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "      gap: 6px;" & vbLf
    s = s & "      font-size: 12px;" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      padding: 3px 5px;" & vbLf
    s = s & "      border-radius: 5px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .col-pick-item:hover { background: rgba(127, 127, 127, 0.12); }" & vbLf
    s = s & "" & vbLf
    s = s & "    .col-pick-actions {" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      gap: 8px;" & vbLf
    s = s & "      margin-bottom: 6px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .col-pick-actions button {" & vbLf
    s = s & "      background: var(--input-bg);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "      border-radius: 6px;" & vbLf
    s = s & "      font-size: 11px;" & vbLf
    s = s & "      padding: 4px 10px;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .col-pick-actions button:hover { border-color: #06b6d4; }" & vbLf
    s = s & "" & vbLf
    s = s & "    /* ---------- Smart Data Table Configuration modal ---------- */" & vbLf
    s = s & "    .step-head {" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "      gap: 8px;" & vbLf
    s = s & "      font-size: 12px;" & vbLf
    s = s & "      font-weight: 800;" & vbLf
    s = s & "      letter-spacing: 0.08em;" & vbLf
    s = s & "      text-transform: uppercase;" & vbLf
    s = s & "      color: #f59e0b;" & vbLf
    s = s & "      margin: 16px 0 10px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .step-badge {" & vbLf
    s = s & "      width: 20px;" & vbLf
    s = s & "      height: 20px;" & vbLf
    s = s & "      border-radius: 50%;" & vbLf
    s = s & "      background: #f59e0b;" & vbLf
    s = s & "      color: #fff;" & vbLf
    s = s & "      font-size: 11px;" & vbLf
    s = s & "      font-weight: 800;" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "      justify-content: center;" & vbLf
    s = s & "      flex-shrink: 0;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .pro-tip {" & vbLf
    s = s & "      border-left: 3px solid #10b981;" & vbLf
    s = s & "      background: rgba(16, 185, 129, 0.10);" & vbLf
    s = s & "      border-radius: 0 6px 6px 0;" & vbLf
    s = s & "      padding: 9px 12px;" & vbLf
    s = s & "      font-size: 11.5px;" & vbLf
    s = s & "      line-height: 1.55;" & vbLf
    s = s & "      color: var(--axis-col);" & vbLf
    s = s & "      margin-bottom: 12px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .pro-tip b { color: #10b981; }" & vbLf
    s = s & "" & vbLf
    s = s & "    .engine-grid {" & vbLf
    s = s & "      display: grid;" & vbLf
    s = s & "      grid-template-columns: 1fr 1fr;" & vbLf
    s = s & "      gap: 12px;" & vbLf
    s = s & "      margin-bottom: 4px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .engine-card {" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      border-radius: 10px;" & vbLf
    s = s & "      background: var(--input-bg);" & vbLf
    s = s & "      padding: 12px;" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      flex-direction: column;" & vbLf
    s = s & "      gap: 12px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .eng-row {" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "      justify-content: space-between;" & vbLf
    s = s & "      gap: 12px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .eng-lab b { display: block; font-size: 12px; color: var(--text-col); }" & vbLf
    s = s & "    .eng-lab span { font-size: 10.5px; color: var(--axis-col); }" & vbLf
    s = s & "    .eng-row .fmt-input { max-width: 155px; flex-shrink: 0; }" & vbLf
    s = s & "" & vbLf
    s = s & "    .switch { position: relative; display: inline-block; width: 42px; height: 22px; flex-shrink: 0; }" & vbLf
    s = s & "    .switch input { opacity: 0; width: 0; height: 0; }" & vbLf
    s = s & "    .switch .knob { position: absolute; left: 0; right: 0; top: 0; bottom: 0; background: rgba(127, 127, 127, 0.35); border-radius: 22px; transition: 0.2s; cursor: pointer; }" & vbLf
    s = s & "    .switch .knob:before { content: """"; position: absolute; width: 16px; height: 16px; left: 3px; top: 3px; background: #fff; border-radius: 50%; transition: 0.2s; }" & vbLf
    s = s & "    .switch input:checked + .knob { background: #10b981; }" & vbLf
    s = s & "    .switch input:checked + .knob:before { transform: translateX(20px); }" & vbLf
    s = s & "" & vbLf
    s = s & "    .lim-row { display: flex; gap: 8px; align-items: center; }" & vbLf
    s = s & "    .lim-row .fmt-input { min-width: 0; }" & vbLf
    s = s & "    .lim-row .lim-dir { max-width: 110px; }" & vbLf
    s = s & "    .lim-row .lim-n { width: 76px; }" & vbLf
    s = s & "    .lim-row .lim-col { flex: 1; }" & vbLf
    s = s & "    .lim-by { font-size: 11px; font-weight: 700; color: var(--axis-col); }" & vbLf
    s = s & "" & vbLf
    s = s & "    .step-sub {" & vbLf
    s = s & "      font-size: 11px;" & vbLf
    s = s & "      font-weight: 700;" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "      border-top: 1px dashed var(--border-col);" & vbLf
    s = s & "      padding-top: 10px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .col-rows {" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      border-radius: 10px;" & vbLf
    s = s & "      background: var(--input-bg);" & vbLf
    s = s & "      overflow: hidden;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .col-rows-head," & vbLf
    s = s & "    .col-row {" & vbLf
    s = s & "      display: grid;" & vbLf
    s = s & "      grid-template-columns: 22px 20px minmax(90px, 1fr) 76px 165px 130px;" & vbLf
    s = s & "      gap: 8px;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "      padding: 7px 10px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .col-rows-head {" & vbLf
    s = s & "      font-size: 10.5px;" & vbLf
    s = s & "      font-weight: 800;" & vbLf
    s = s & "      letter-spacing: 0.05em;" & vbLf
    s = s & "      text-transform: uppercase;" & vbLf
    s = s & "      color: var(--axis-col);" & vbLf
    s = s & "      background: var(--card-bg);" & vbLf
    s = s & "      border-bottom: 1px solid var(--border-col);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .col-row {" & vbLf
    s = s & "      border-bottom: 1px solid var(--border-col);" & vbLf
    s = s & "      background: var(--card-bg);" & vbLf
    s = s & "      font-size: 12px;" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .col-row:last-child { border-bottom: none; }" & vbLf
    s = s & "    .col-row:hover { background: rgba(127, 127, 127, 0.07); }" & vbLf
    s = s & "    .col-row.drag-over { box-shadow: inset 0 2px 0 #f59e0b; }" & vbLf
    s = s & "    .col-row.off { opacity: 0.5; }" & vbLf
    s = s & "" & vbLf
    s = s & "    .drag-h { cursor: grab; color: var(--axis-col); font-size: 14px; text-align: center; user-select: none; }" & vbLf
    s = s & "    .drag-h:active { cursor: grabbing; }" & vbLf
    s = s & "" & vbLf
    s = s & "    .col-name," & vbLf
    s = s & "    .col-calc," & vbLf
    s = s & "    .col-sdate {" & vbLf
    s = s & "      width: 100%;" & vbLf
    s = s & "      box-sizing: border-box;" & vbLf
    s = s & "      background: var(--input-bg);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "      border-radius: 6px;" & vbLf
    s = s & "      padding: 5px 7px;" & vbLf
    s = s & "      font-size: 12px;" & vbLf
    s = s & "      font-family: inherit;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .col-src {" & vbLf
    s = s & "      font-size: 10px;" & vbLf
    s = s & "      color: var(--axis-col);" & vbLf
    s = s & ""
    Sec01_HeadCss_1 = s
End Function

'--- Sec01_HeadCss_2 : Document head, CDN libraries (ECharts/Three.js/GridStack/XLSX) and the complete CSS theme (chunk 2 of 3)
Private Function Sec01_HeadCss_2() As String
    Dim s As String
    s = s & "      text-align: center;" & vbLf
    s = s & "      background: var(--input-bg);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      border-radius: 5px;" & vbLf
    s = s & "      padding: 4px 4px;" & vbLf
    s = s & "      white-space: nowrap;" & vbLf
    s = s & "      overflow: hidden;" & vbLf
    s = s & "      text-overflow: ellipsis;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .col-row input[type=""checkbox""] { accent-color: #10b981; }" & vbLf
    s = s & "" & vbLf
    s = s & "    .btn-tbl-generate {" & vbLf
    s = s & "      background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%);" & vbLf
    s = s & "      color: #fff;" & vbLf
    s = s & "      border: none;" & vbLf
    s = s & "      border-radius: 8px;" & vbLf
    s = s & "      padding: 9px 18px;" & vbLf
    s = s & "      font-size: 13px;" & vbLf
    s = s & "      font-weight: 700;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      box-shadow: 0 3px 10px rgba(245, 158, 11, 0.35);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .btn-tbl-generate:hover { filter: brightness(1.07); }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dash-table.wrap td {" & vbLf
    s = s & "      white-space: normal;" & vbLf
    s = s & "      overflow: visible;" & vbLf
    s = s & "      text-overflow: clip;" & vbLf
    s = s & "      word-break: break-word;" & vbLf
    s = s & "      max-width: 260px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .chk-row {" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "      gap: 8px;" & vbLf
    s = s & "      font-size: 12px;" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "      margin-top: 4px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .slicer-container {" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      gap: 15px;" & vbLf
    s = s & "      flex-wrap: wrap;" & vbLf
    s = s & "      padding: 20px 30px 0 30px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .slicer-card {" & vbLf
    s = s & "      position: relative;" & vbLf
    s = s & "      background: var(--card-bg);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      border-radius: 16px;" & vbLf
    s = s & "      padding: 20px;" & vbLf
    s = s & "      box-sizing: border-box;" & vbLf
    s = s & "      width: 100%;" & vbLf
    s = s & "      height: 100%;" & vbLf
    s = s & "      overflow: hidden;" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      flex-direction: column;" & vbLf
    s = s & "      align-items: flex-start;" & vbLf
    s = s & "      justify-content: center;" & vbLf
    s = s & "      box-shadow: 0 4px 6px -1px var(--shadow-col);" & vbLf
    s = s & "      transition: all 0.3s;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .slicer-card:hover {" & vbLf
    s = s & "      border-color: #f43f5e;" & vbLf
    s = s & "      box-shadow: 0 10px 15px -3px var(--shadow-col);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .slicer-info {" & vbLf
    s = s & "      position: relative;" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      flex-direction: column;" & vbLf
    s = s & "      gap: 4px;" & vbLf
    s = s & "      width: 100%;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .slicer-label {" & vbLf
    s = s & "      font-size: 10px;" & vbLf
    s = s & "      font-weight: 800;" & vbLf
    s = s & "      color: #64748b;" & vbLf
    s = s & "      text-transform: uppercase;" & vbLf
    s = s & "      letter-spacing: 0.5px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .slicer-select {" & vbLf
    s = s & "      background: var(--input-bg);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "      font-size: 13px;" & vbLf
    s = s & "      font-weight: 600;" & vbLf
    s = s & "      font-family: 'Inter', sans-serif;" & vbLf
    s = s & "      padding: 6px 10px;" & vbLf
    s = s & "      border-radius: 6px;" & vbLf
    s = s & "      outline: none;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      min-width: 150px;" & vbLf
    s = s & "      transition: 0.2s;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .slicer-select:focus {" & vbLf
    s = s & "      border-color: #f43f5e;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .delete-slicer {" & vbLf
    s = s & "      position: absolute;" & vbLf
    s = s & "      top: 8px;" & vbLf
    s = s & "      right: 8px;" & vbLf
    s = s & "      background: transparent;" & vbLf
    s = s & "      border: none;" & vbLf
    s = s & "      color: #ef4444;" & vbLf
    s = s & "      font-size: 18px;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      opacity: 0;" & vbLf
    s = s & "      transition: opacity 0.2s;" & vbLf
    s = s & "      z-index: 10;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .slicer-card:hover .delete-slicer {" & vbLf
    s = s & "      opacity: 1;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .kpi-container {" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      gap: 20px;" & vbLf
    s = s & "      flex-wrap: wrap;" & vbLf
    s = s & "      padding: 25px 30px 10px 30px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .kpi-card {" & vbLf
    s = s & "      position: relative;" & vbLf
    s = s & "      flex: 1;" & vbLf
    s = s & "      min-width: 180px;" & vbLf
    s = s & "      background-color: var(--card-bg);" & vbLf
    s = s & "      border-radius: 16px;" & vbLf
    s = s & "      box-shadow: 0 4px 6px -1px var(--shadow-col);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      padding: 20px;" & vbLf
    s = s & "      text-align: center;" & vbLf
    s = s & "      transition: all 0.3s ease;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .kpi-card:hover {" & vbLf
    s = s & "      transform: translateY(-4px);" & vbLf
    s = s & "      box-shadow: 0 10px 15px -3px var(--shadow-col);" & vbLf
    s = s & "      border-color: #6366f1;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .kpi-title {" & vbLf
    s = s & "      font-size: 13px;" & vbLf
    s = s & "      color: var(--kpi-title-col);" & vbLf
    s = s & "      text-transform: uppercase;" & vbLf
    s = s & "      font-weight: 600;" & vbLf
    s = s & "      letter-spacing: 0.5px;" & vbLf
    s = s & "      margin-bottom: 8px;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .kpi-val {" & vbLf
    s = s & "      font-size: 28px;" & vbLf
    s = s & "      font-weight: 700;" & vbLf
    s = s & "      color: var(--kpi-val-col);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .delete-kpi," & vbLf
    s = s & "    .delete-chart {" & vbLf
    s = s & "      position: absolute;" & vbLf
    s = s & "      top: 8px;" & vbLf
    s = s & "      right: 8px;" & vbLf
    s = s & "      background: transparent;" & vbLf
    s = s & "      border: none;" & vbLf
    s = s & "      color: #ef4444;" & vbLf
    s = s & "      font-size: 18px;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      opacity: 0;" & vbLf
    s = s & "      transition: opacity 0.2s;" & vbLf
    s = s & "      z-index: 10;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .kpi-card:hover .delete-kpi," & vbLf
    s = s & "    .chart-card:hover .delete-chart {" & vbLf
    s = s & "      opacity: 1;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .chart-grid {" & vbLf
    s = s & "      display: grid;" & vbLf
    s = s & "      grid-template-columns: repeat(2, 1fr);" & vbLf
    s = s & "      gap: 25px;" & vbLf
    s = s & "      padding: 15px 30px 30px 30px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .chart-card {" & vbLf
    s = s & "      position: relative;" & vbLf
    s = s & "      background-color: var(--card-bg);" & vbLf
    s = s & "      border-radius: 16px;" & vbLf
    s = s & "      box-shadow: 0 4px 6px -1px var(--shadow-col);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      flex-direction: column;" & vbLf
    s = s & "      overflow: hidden;" & vbLf
    s = s & "      transition: all 0.3s ease;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .chart-card:hover {" & vbLf
    s = s & "      transform: translateY(-4px);" & vbLf
    s = s & "      box-shadow: 0 10px 15px -3px var(--shadow-col);" & vbLf
    s = s & "      border-color: var(--kpi-val-col);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .chart-header {" & vbLf
    s = s & "      padding: 18px 20px 0px 20px;" & vbLf
    s = s & "      font-weight: 600;" & vbLf
    s = s & "      font-size: 16px;" & vbLf
    s = s & "      color: var(--chart-header-col);" & vbLf
    s = s & "      text-align: center;" & vbLf
    s = s & "      height: 45px;" & vbLf
    s = s & "      white-space: nowrap;" & vbLf
    s = s & "      overflow: hidden;" & vbLf
    s = s & "      text-overflow: ellipsis;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .fullscreen {" & vbLf
    s = s & "      position: fixed !important;" & vbLf
    s = s & "      top: 0;" & vbLf
    s = s & "      left: 0;" & vbLf
    s = s & "      width: 100vw !important;" & vbLf
    s = s & "      height: 100vh !important;" & vbLf
    s = s & "      z-index: 9999;" & vbLf
    s = s & "      background: var(--card-bg);" & vbLf
    s = s & "      padding: 20px;" & vbLf
    s = s & "      border-radius: 0 !important;" & vbLf
    s = s & "      margin: 0 !important;" & vbLf
    s = s & "      transform: none !important;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .modal-overlay {" & vbLf
    s = s & "      display: none;" & vbLf
    s = s & "      position: fixed;" & vbLf
    s = s & "      top: 0;" & vbLf
    s = s & "      left: 0;" & vbLf
    s = s & "      width: 100%;" & vbLf
    s = s & "      height: 100%;" & vbLf
    s = s & "      background: rgba(0, 0, 0, 0.6);" & vbLf
    s = s & "      backdrop-filter: blur(4px);" & vbLf
    s = s & "      z-index: 10000;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .modal-box {" & vbLf
    s = s & "      position: absolute;" & vbLf
    s = s & "      top: 50%;" & vbLf
    s = s & "      left: 50%;" & vbLf
    s = s & "      transform: translate(-50%, -50%);" & vbLf
    s = s & "      width: 400px;" & vbLf
    s = s & "      max-width: 95%;" & vbLf
    s = s & "      max-height: 85vh;" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      flex-direction: column;" & vbLf
    s = s & "      background: var(--card-bg);" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "      padding: 25px;" & vbLf
    s = s & "      border-radius: 16px;" & vbLf
    s = s & "      box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .modal-wide {" & vbLf
    s = s & "      width: 650px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .modal-title {" & vbLf
    s = s & "      font-size: 20px;" & vbLf
    s = s & "      font-weight: 700;" & vbLf
    s = s & "      margin-bottom: 20px;" & vbLf
    s = s & "      padding-bottom: 12px;" & vbLf
    s = s & "      border-bottom: 1px solid var(--border-col);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .modal-body {" & vbLf
    s = s & "      overflow-y: auto;" & vbLf
    s = s & "      padding-right: 10px;" & vbLf
    s = s & "      margin-bottom: 10px;" & vbLf
    s = s & "      flex-grow: 1;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .modal-body::-webkit-scrollbar {" & vbLf
    s = s & "      width: 6px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .modal-body::-webkit-scrollbar-thumb {" & vbLf
    s = s & "      background: var(--scroll-thumb);" & vbLf
    s = s & "      border-radius: 4px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .modal-footer {" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      justify-content: flex-end;" & vbLf
    s = s & "      gap: 12px;" & vbLf
    s = s & "      padding-top: 15px;" & vbLf
    s = s & "      border-top: 1px solid var(--border-col);" & vbLf
    s = s & "      margin-top: auto;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .fmt-row {" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      gap: 15px;" & vbLf
    s = s & "      margin-bottom: 15px;" & vbLf
    s = s & "      width: 100%;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .input-group {" & vbLf
    s = s & "      flex: 1;" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      flex-direction: column;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .input-label {" & vbLf
    s = s & "      font-size: 11px;" & vbLf
    s = s & "      font-weight: 700;" & vbLf
    s = s & "      color: #64748b;" & vbLf
    s = s & "      margin-bottom: 6px;" & vbLf
    s = s & "      text-transform: uppercase;" & vbLf
    s = s & "      letter-spacing: 0.5px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    select.fmt-input," & vbLf
    s = s & "    input.fmt-input {" & vbLf
    s = s & "      width: 100%;" & vbLf
    s = s & "      padding: 10px 14px;" & vbLf
    s = s & "      background: var(--input-bg);" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      border-radius: 8px;" & vbLf
    s = s & "      font-family: 'Inter', sans-serif;" & vbLf
    s = s & "      font-size: 13px;" & vbLf
    s = s & "      transition: all 0.2s;" & vbLf
    s = s & "      outline: none;" & vbLf
    s = s & "      box-sizing: border-box;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    select.fmt-input:focus," & vbLf
    s = s & "    input.fmt-input:focus {" & vbLf
    s = s & "      border-color: #6366f1;" & vbLf
    s = s & "      box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.15);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .btn-primary {" & vbLf
    s = s & "      background: linear-gradient(135deg, #10b981 0%, #059669 100%);" & vbLf
    s = s & "      color: white;" & vbLf
    s = s & "      border: none;" & vbLf
    s = s & "      padding: 10px 24px;" & vbLf
    s = s & "      border-radius: 8px;" & vbLf
    s = s & "      font-weight: 600;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      transition: all 0.2s;" & vbLf
    s = s & "      font-size: 13px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .btn-primary:hover {" & vbLf
    s = s & "      transform: translateY(-1px);" & vbLf
    s = s & "      box-shadow: 0 4px 12px rgba(16, 185, 129, 0.3);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .btn-secondary {" & vbLf
    s = s & "      background: transparent;" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      padding: 10px 24px;" & vbLf
    s = s & "      border-radius: 8px;" & vbLf
    s = s & "      font-weight: 600;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      transition: all 0.2s;" & vbLf
    s = s & "      font-size: 13px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .btn-secondary:hover {" & vbLf
    s = s & "      background: var(--border-col);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .fmt-btn {" & vbLf
    s = s & "      display: block;" & vbLf
    s = s & "      width: 100%;" & vbLf
    s = s & "      padding: 10px;" & vbLf
    s = s & "      margin: 8px 0;" & vbLf
    s = s & "      background: var(--input-bg);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      border-radius: 8px;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "      font-size: 13px;" & vbLf
    s = s & "      font-weight: 500;" & vbLf
    s = s & "      transition: 0.2s;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .fmt-btn:hover {" & vbLf
    s = s & "      background: var(--border-col);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .grid-stack {" & vbLf
    s = s & "      background: transparent;" & vbLf
    s = s & "      padding: 15px;" & vbLf
    s = s & "      min-height: 500px;" & vbLf
    s = s & "      flex-grow: 1;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .grid-stack-item-content {" & vbLf
    s = s & "      border-radius: 12px;" & vbLf
    s = s & "      background: transparent;" & vbLf
    s = s & "      overflow: hidden !important;" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      flex-direction: column;" & vbLf
    s = s & "      transition: border 0.2s;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .placeholder-card {" & vbLf
    s = s & "      width: 100%;" & vbLf
    s = s & "      height: 100%;" & vbLf
    s = s & "      border: 2px dashed var(--border-col);" & vbLf
    s = s & "      border-radius: 12px;" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "      justify-content: center;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      transition: 0.3s;" & vbLf
    s = s & "      color: var(--axis-col);" & vbLf
    s = s & "      font-weight: 600;" & vbLf
    s = s & "      background: var(--card-bg);" & vbLf
    s = s & "      font-size: 14px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .placeholder-card:hover {" & vbLf
    s = s & "      border-color: #6366f1;" & vbLf
    s = s & "      background: rgba(99, 102, 241, 0.05);" & vbLf
    s = s & "      color: #6366f1;" & vbLf
    s = s & "      transform: scale(0.98);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .layout-grid {" & vbLf
    s = s & "      display: grid;" & vbLf
    s = s & "      grid-template-columns: repeat(2, 1fr);" & vbLf
    s = s & "      gap: 15px;" & vbLf
    s = s & "      padding: 10px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .layout-option {" & vbLf
    s = s & "      border: 2px solid var(--border-col);" & vbLf
    s = s & "      border-radius: 8px;" & vbLf
    s = s & "      padding: 10px;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      transition: 0.2s;" & vbLf
    s = s & "      background: var(--card-bg);" & vbLf
    s = s & "      text-align: center;" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .layout-option:hover {" & vbLf
    s = s & "      border-color: #4f46e5;" & vbLf
    s = s & "      box-shadow: 0 4px 12px rgba(79, 70, 229, 0.2);" & vbLf
    s = s & "      transform: translateY(-2px);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .edit-toggle-wrapper {" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "      gap: 8px;" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "      font-weight: 600;" & vbLf
    s = s & "      font-size: 13px;" & vbLf
    s = s & "      margin-right: 15px;" & vbLf
    s = s & "      border-right: 1px solid var(--border-col);" & vbLf
    s = s & "      padding-right: 15px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .switch {" & vbLf
    s = s & "      position: relative;" & vbLf
    s = s & "      display: inline-block;" & vbLf
    s = s & "      width: 36px;" & vbLf
    s = s & "      height: 18px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .switch input {" & vbLf
    s = s & "      opacity: 0;" & vbLf
    s = s & "      width: 0;" & vbLf
    s = s & "      height: 0;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .slider {" & vbLf
    s = s & "      position: absolute;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      top: 0;" & vbLf
    s = s & "      left: 0;" & vbLf
    s = s & "      right: 0;" & vbLf
    s = s & "      bottom: 0;" & vbLf
    s = s & "      background-color: #cbd5e1;" & vbLf
    s = s & "      transition: .4s;" & vbLf
    s = s & "      border-radius: 20px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .slider:before {" & vbLf
    s = s & "      position: absolute;" & vbLf
    s = s & "      content: '';" & vbLf
    s = s & "      height: 12px;" & vbLf
    s = s & "      width: 12px;" & vbLf
    s = s & "      left: 3px;" & vbLf
    s = s & "      bottom: 3px;" & vbLf
    s = s & "      background-color: white;" & vbLf
    s = s & "      transition: .4s;" & vbLf
    s = s & "      border-radius: 50%;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    input:checked+.slider {" & vbLf
    s = s & "      background-color: #10b981;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    input:checked+.slider:before {" & vbLf
    s = s & "      transform: translateX(18px);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .kpi-card {" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      flex-direction: column;" & vbLf
    s = s & "      justify-content: center;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "      margin: 0 !important;" & vbLf
    s = s & "      height: 100% !important;" & vbLf
    s = s & "      width: 100% !important;" & vbLf
    s = s & "      box-sizing: border-box;" & vbLf
    s = s & "      padding: 15px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .chart-card {" & vbLf
    s = s & "      margin: 0 !important;" & vbLf
    s = s & "      height: 100% !important;" & vbLf
    s = s & "      width: 100% !important;" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      flex-direction: column;" & vbLf
    s = s & "      box-sizing: border-box;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .layout-sections {" & vbLf
    s = s & "      display: grid;" & vbLf
    s = s & "      grid-template-columns: minmax(260px, 1fr) minmax(320px, 1.35fr);" & vbLf
    s = s & "      gap: 20px;" & vbLf
    s = s & "      align-items: start;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    @media (max-width: 720px) {" & vbLf
    s = s & "      .layout-sections { grid-template-columns: 1fr; }" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .custom-builder-card {" & vbLf
    s = s & "      background: var(--card-bg);" & vbLf
    s = s & "      border: 1px solid #8b5cf6;" & vbLf
    s = s & "      padding: 15px;" & vbLf
    s = s & "      border-radius: 12px;" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      flex-wrap: wrap;" & vbLf
    s = s & "      gap: 15px;" & vbLf
    s = s & "      align-items: flex-end;" & vbLf
    s = s & "      box-shadow: 0 4px 15px rgba(139, 92, 246, 0.15);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .custom-builder-card .input-group {" & vbLf
    s = s & "      flex: 1;" & vbLf
    s = s & "      margin-bottom: 0 !important;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .btn-generate {" & vbLf
    s = s & "      background: linear-gradient(135deg, #8b5cf6 0%, #6d28d9 100%);" & vbLf
    s = s & "      color: white;" & vbLf
    s = s & "      border: none;" & vbLf
    s = s & "      padding: 10px 15px;" & vbLf
    s = s & "      border-radius: 8px;" & vbLf
    s = s & "      font-weight: 600;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      height: 38px;" & vbLf
    s = s & "      transition: 0.2s;" & vbLf
    s = s & "      white-space: nowrap;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .btn-generate:hover {" & vbLf
    s = s & "      transform: translateY(-2px);" & vbLf
    s = s & "      box-shadow: 0 6px 15px rgba(139, 92, 246, 0.3);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .slicer-placeholder {" & vbLf
    s = s & "      border: 2px dashed #f43f5e;" & vbLf
    s = s & "      color: #f43f5e;" & vbLf
    s = s & "      background: transparent;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "      justify-content: center;" & vbLf
    s = s & "      border-radius: 16px;" & vbLf
    s = s & "      box-sizing: border-box;" & vbLf
    s = s & "      width: 100%;" & vbLf
    s = s & "      height: 100%;" & vbLf
    s = s & "      padding: 10px;" & vbLf
    s = s & "      font-size: 13px;" & vbLf
    s = s & "      font-weight: 600;" & vbLf
    s = s & "      text-align: center;" & vbLf
    s = s & "      transition: 0.2s;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .slicer-placeholder:hover {" & vbLf
    s = s & "      background: rgba(244, 63, 94, 0.05);" & vbLf
    s = s & "      transform: scale(0.98);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .slicer-btn {" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "      justify-content: space-between;" & vbLf
    s = s & "      gap: 8px;" & vbLf
    s = s & "      width: 100%;" & vbLf
    s = s & "      text-align: left;" & vbLf
    s = s & "      background: var(--input-bg);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "      font-size: 13px;" & vbLf
    s = s & "      font-weight: 600;" & vbLf
    s = s & "      padding: 7px 10px;" & vbLf
    s = s & "      border-radius: 6px;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      font-family: 'Inter', sans-serif;" & vbLf
    s = s & "      transition: 0.2s;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .slicer-btn:hover { border-color: #f43f5e; }" & vbLf
    s = s & "    .slicer-caret { margin-left: auto; font-size: 10px; opacity: 0.7; }" & vbLf
    s = s & "" & vbLf
    s = s & "    .slicer-pop {" & vbLf
    s = s & "      display: none;" & vbLf
    s = s & "      position: absolute;" & vbLf
    s = s & "      left: 0;" & vbLf
    s = s & "      right: 0;" & vbLf
    s = s & "      top: 100%;" & vbLf
    s = s & "      margin-top: 6px;" & vbLf
    s = s & "      background: var(--card-bg);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      border-radius: 10px;" & vbLf
    s = s & "      box-shadow: 0 14px 28px rgba(0, 0, 0, 0.28);" & vbLf
    s = s & "      max-height: 250px;" & vbLf
    s = s & "      overflow-y: auto;" & vbLf
    s = s & "      z-index: 80;" & vbLf
    s = s & "      padding: 4px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .slicer-pop-item {" & vbLf
    s = s & "      padding: 7px 10px;" & vbLf
    s = s & "      font-size: 12.5px;" & vbLf
    s = s & "      font-weight: 600;" & vbLf
    s = s & "      border-radius: 6px;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "      white-space: nowrap;" & vbLf
    s = s & "      overflow: hidden;" & vbLf
    s = s & "      text-overflow: ellipsis;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .slicer-pop-item:hover { background: rgba(244, 63, 94, 0.12); color: #f43f5e; }" & vbLf
    s = s & "    .slicer-pop-item.on { background: rgba(244, 63, 94, 0.18); color: #f43f5e; }" & vbLf
    s = s & "" & vbLf
    s = s & "    .curr-badge {" & vbLf
    s = s & "      display: inline-flex;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "      justify-content: center;" & vbLf
    s = s & "      min-width: 26px;" & vbLf
    s = s & "      background: var(--input-bg);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      color: var(--kpi-val-col);" & vbLf
    s = s & "      font-size: 13px;" & vbLf
    s = s & "      font-weight: 800;" & vbLf
    s = s & "      padding: 3px 10px;" & vbLf
    s = s & "      border-radius: 999px;" & vbLf
    s = s & "      margin-left: 10px;" & vbLf
    s = s & "      letter-spacing: 0.4px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .size-row { display: flex; gap: 15px; }" & vbLf
    s = s & "" & vbLf
    s = s & "    .chart-body {" & vbLf
    s = s & "      flex: 1;" & vbLf
    s = s & "      width: 100%;" & vbLf
    s = s & "      height: 100%;" & vbLf
    s = s & "      min-height: 300px;" & vbLf
    s = s & "      padding: 10px;" & vbLf
    s = s & "      position: relative;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .chart-box {" & vbLf
    s = s & "      position: absolute;" & vbLf
    s = s & "      top: 0;" & vbLf
    s = s & "      left: 0;" & vbLf
    s = s & "      right: 0;" & vbLf
    s = s & "      bottom: 0;" & vbLf
    s = s & "      width: 100% !important;" & vbLf
    s = s & "      height: 100% !important;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dash-footer {" & vbLf
    s = s & "      position: fixed;" & vbLf
    s = s & "      left: 0;" & vbLf
    s = s & "      right: 0;" & vbLf
    s = s & "      bottom: 0;" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      justify-content: space-between;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "      padding: 15px 30px;" & vbLf
    s = s & "      background: var(--header-bg);" & vbLf
    s = s & "      backdrop-filter: blur(12px);" & vbLf
    s = s & "      border-top: 1px solid var(--border-col);" & vbLf
    s = s & "      color: var(--axis-col);" & vbLf
    s = s & "      font-size: 13px;" & vbLf
    s = s & "      font-weight: 500;" & vbLf
    s = s & "      z-index: 70;" & vbLf
    s = s & "      transition: all 0.3s;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .footer-left {" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "      gap: 15px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .footer-label {" & vbLf
    s = s & "      font-weight: 700;" & vbLf
    s = s & "      text-transform: uppercase;" & vbLf
    s = s & "      letter-spacing: 0.5px;" & vbLf
    s = s & "      font-size: 11px;" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "      opacity: 0.7;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .footer-btn {" & vbLf
    s = s & "      display: inline-flex;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "      gap: 6px;" & vbLf
    s = s & "      padding: 6px 11px;" & vbLf
    s = s & "      font-size: 11px;" & vbLf
    s = s & "      font-weight: 700;" & vbLf
    s = s & "      letter-spacing: 0.3px;" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "      background: var(--input-bg);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      border-radius: 8px;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      transition: all 0.2s;" & vbLf
    s = s & "      white-space: nowrap;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .footer-btn:hover {" & vbLf
    s = s & "      border-color: var(--kpi-val-col);" & vbLf
    s = s & "      color: var(--kpi-val-col);" & vbLf
    s = s & "      transform: translateY(-1px);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .footer-btn i {" & vbLf
    s = s & "      font-size: 12px;" & vbLf
    s = s & "      color: var(--kpi-val-col);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .theme-switcher {" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      gap: 10px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .theme-swatch {" & vbLf
    s = s & "      width: 22px;" & vbLf
    s = s & "      height: 22px;" & vbLf
    s = s & "      border-radius: 50%;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      border: 2px solid transparent;" & vbLf
    s = s & "      transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);" & vbLf
    s = s & "      box-shadow: 0 2px 5px rgba(0, 0, 0, 0.15);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .theme-swatch:hover {" & vbLf
    s = s & "      transform: scale(1.15) translateY(-2px);" & vbLf
    s = s & "      box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .theme-swatch.active {" & vbLf
    s = s & "      border-color: var(--kpi-val-col);" & vbLf
    s = s & "      transform: scale(1.15);" & vbLf
    s = s & "      box-shadow: 0 0 0 2px var(--body-bg), 0 0 0 4px var(--kpi-val-col);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .sw-light {" & vbLf
    s = s & "      background: #f8fafc;" & vbLf
    s = s & "      border: 1px solid #cbd5e1;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .sw-dark {" & vbLf
    s = s & "      background: #09090b;" & vbLf
    s = s & "      border: 1px solid #3f3f46;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .sw-ocean {" & vbLf
    s = s & "      background: #e0f2fe;" & vbLf
    s = s & "      border: 1px solid #7dd3fc;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .sw-midnight {" & vbLf
    s = s & "      background: #020617;" & vbLf
    s = s & "      border: 1px solid #334155;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .sw-animated {" & vbLf
    s = s & "      background: linear-gradient(135deg, #00f3ff, #bc13fe);" & vbLf
    s = s & "      border: 1px solid #ffffff;" & vbLf
    s = s & ""
    Sec01_HeadCss_2 = s
End Function

'--- Sec01_HeadCss_3 : Document head, CDN libraries (ECharts/Three.js/GridStack/XLSX) and the complete CSS theme (chunk 3 of 3)
Private Function Sec01_HeadCss_3() As String
    Dim s As String
    s = s & "      animation: gradientBG 3s ease infinite;" & vbLf
    s = s & "      background-size: 200% 200%;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .sw-forest {" & vbLf
    s = s & "      background: linear-gradient(135deg, #16a34a, #86efac);" & vbLf
    s = s & "      border: 1px solid #166534;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .sw-rose {" & vbLf
    s = s & "      background: linear-gradient(135deg, #e11d48, #fda4af);" & vbLf
    s = s & "      border: 1px solid #9f1239;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .sw-navy {" & vbLf
    s = s & "      background: linear-gradient(135deg, #0b1a33, #3b82f6);" & vbLf
    s = s & "      border: 1px solid #60a5fa;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .footer-right {" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "      gap: 6px;" & vbLf
    s = s & "      letter-spacing: 0.2px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .footer-brand {" & vbLf
    s = s & "      font-weight: 800;" & vbLf
    s = s & "      color: var(--kpi-val-col);" & vbLf
    s = s & "      background: linear-gradient(135deg, var(--kpi-val-col) 0%, #a855f7 100%);" & vbLf
    s = s & "      -webkit-background-clip: text;" & vbLf
    s = s & "      -webkit-text-fill-color: transparent;" & vbLf
    s = s & "      font-size: 14px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dv-tabs {" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      flex-wrap: wrap;" & vbLf
    s = s & "      gap: 8px;" & vbLf
    s = s & "      margin-bottom: 12px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dv-tab {" & vbLf
    s = s & "      padding: 6px 14px;" & vbLf
    s = s & "      font-size: 12px;" & vbLf
    s = s & "      font-weight: 700;" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "      background: var(--input-bg);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      border-radius: 8px;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      transition: all 0.2s;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dv-tab:hover {" & vbLf
    s = s & "      border-color: var(--kpi-val-col);" & vbLf
    s = s & "      color: var(--kpi-val-col);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dv-tab.on {" & vbLf
    s = s & "      color: #fff;" & vbLf
    s = s & "      background: linear-gradient(135deg, var(--kpi-val-col) 0%, #a855f7 100%);" & vbLf
    s = s & "      border-color: transparent;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dv-wrap {" & vbLf
    s = s & "      max-height: 440px;" & vbLf
    s = s & "      overflow: auto;" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      border-radius: 10px;" & vbLf
    s = s & "      background: var(--card-bg);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dv-wrap .dash-table th {" & vbLf
    s = s & "      background: var(--card-bg);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dm-hint {" & vbLf
    s = s & "      font-size: 12px;" & vbLf
    s = s & "      color: var(--axis-col);" & vbLf
    s = s & "      background: var(--input-bg);" & vbLf
    s = s & "      border: 1px dashed var(--border-col);" & vbLf
    s = s & "      border-radius: 8px;" & vbLf
    s = s & "      padding: 8px 10px;" & vbLf
    s = s & "      margin-bottom: 10px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dm-canvas {" & vbLf
    s = s & "      position: relative;" & vbLf
    s = s & "      height: 420px;" & vbLf
    s = s & "      overflow: auto;" & vbLf
    s = s & "      background:" & vbLf
    s = s & "        repeating-linear-gradient(0deg, transparent, transparent 19px, rgba(127, 127, 127, 0.07) 20px)," & vbLf
    s = s & "        repeating-linear-gradient(90deg, transparent, transparent 19px, rgba(127, 127, 127, 0.07) 20px)," & vbLf
    s = s & "        var(--card-bg);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      border-radius: 10px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dm-svg {" & vbLf
    s = s & "      position: absolute;" & vbLf
    s = s & "      left: 0;" & vbLf
    s = s & "      top: 0;" & vbLf
    s = s & "      width: 100%;" & vbLf
    s = s & "      height: 100%;" & vbLf
    s = s & "      pointer-events: none;" & vbLf
    s = s & "      z-index: 1;" & vbLf
    s = s & "      overflow: visible;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dm-table {" & vbLf
    s = s & "      position: absolute;" & vbLf
    s = s & "      width: 172px;" & vbLf
    s = s & "      background: var(--card-bg);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      border-radius: 10px;" & vbLf
    s = s & "      box-shadow: 0 6px 14px -6px var(--shadow-col);" & vbLf
    s = s & "      z-index: 2;" & vbLf
    s = s & "      overflow: hidden;" & vbLf
    s = s & "      user-select: none;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dm-head {" & vbLf
    s = s & "      padding: 8px 10px;" & vbLf
    s = s & "      font-size: 11px;" & vbLf
    s = s & "      font-weight: 800;" & vbLf
    s = s & "      text-transform: uppercase;" & vbLf
    s = s & "      letter-spacing: 0.4px;" & vbLf
    s = s & "      color: #fff;" & vbLf
    s = s & "      background: linear-gradient(135deg, var(--kpi-val-col) 0%, #a855f7 100%);" & vbLf
    s = s & "      cursor: move;" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      justify-content: space-between;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "      gap: 6px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dm-head span {" & vbLf
    s = s & "      overflow: hidden;" & vbLf
    s = s & "      text-overflow: ellipsis;" & vbLf
    s = s & "      white-space: nowrap;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dm-col {" & vbLf
    s = s & "      padding: 6px 10px;" & vbLf
    s = s & "      font-size: 12px;" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "      border-top: 1px solid var(--border-col);" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      white-space: nowrap;" & vbLf
    s = s & "      overflow: hidden;" & vbLf
    s = s & "      text-overflow: ellipsis;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dm-col:hover {" & vbLf
    s = s & "      background: var(--input-bg);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dm-col.on {" & vbLf
    s = s & "      background: var(--kpi-val-col);" & vbLf
    s = s & "      color: #fff;" & vbLf
    s = s & "      font-weight: 700;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dm-rels {" & vbLf
    s = s & "      margin-top: 10px;" & vbLf
    s = s & "      display: flex;" & vbLf
    s = s & "      flex-wrap: wrap;" & vbLf
    s = s & "      gap: 8px;" & vbLf
    s = s & "      min-height: 26px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dm-rel {" & vbLf
    s = s & "      display: inline-flex;" & vbLf
    s = s & "      align-items: center;" & vbLf
    s = s & "      gap: 8px;" & vbLf
    s = s & "      padding: 5px 9px;" & vbLf
    s = s & "      font-size: 11px;" & vbLf
    s = s & "      font-weight: 700;" & vbLf
    s = s & "      color: var(--text-col);" & vbLf
    s = s & "      background: var(--input-bg);" & vbLf
    s = s & "      border: 1px solid var(--border-col);" & vbLf
    s = s & "      border-radius: 999px;" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dm-rel b {" & vbLf
    s = s & "      color: var(--kpi-val-col);" & vbLf
    s = s & "    }" & vbLf
    s = s & "" & vbLf
    s = s & "    .dm-rel button {" & vbLf
    s = s & "      border: none;" & vbLf
    s = s & "      background: transparent;" & vbLf
    s = s & "      color: #ef4444;" & vbLf
    s = s & "      font-size: 13px;" & vbLf
    s = s & "      cursor: pointer;" & vbLf
    s = s & "      line-height: 1;" & vbLf
    s = s & "      padding: 0;" & vbLf
    s = s & "    }" & vbLf
    s = s & "  </style>" & vbLf
    s = s & "</head>" & vbLf
    s = s & "" & vbLf
    s = s & "<body class=""animated-theme"">" & vbLf
    s = s & "  <style type=""text/css"" gs-style-id=""gs-id-0""></style>" & vbLf
    s = s & "  "
    Sec01_HeadCss_3 = s
End Function

'--- Sec02_BasicModals : Add Slicer, Update Data (Excel) and Label Settings modal markup
Private Function Sec02_BasicModals() As String
    Dim s As String
    s = s & "<div id=""addSlicerModal"" class=""modal-overlay"" style=""display: none;"">" & vbLf
    s = s & "    <div class=""modal-box"">" & vbLf
    s = s & "      <div class=""modal-title"">Add Advanced Slicer</div>" & vbLf
    s = s & "      <div class=""modal-body"">" & vbLf
    s = s & "        <div class=""input-group"" style=""margin-bottom:15px;"">" & vbLf
    s = s & "          <label class=""input-label"">Filter Column</label>" & vbLf
    s = s & "          <select id=""newSlicerCol"" class=""fmt-input""></select>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "        <div id=""slicerDateDiv"" class=""input-group"" style=""margin-bottom:15px; display:none;"">" & vbLf
    s = s & "          <label class=""input-label"">Date Hierarchy</label>" & vbLf
    s = s & "          <select id=""newSlicerHierarchy"" class=""fmt-input"">" & vbLf
    s = s & "            <option value=""None"">None</option>" & vbLf
    s = s & "            <option value=""Year"">Year</option>" & vbLf
    s = s & "            <option value=""Quarter"">Quarter</option>" & vbLf
    s = s & "            <option value=""Month"">Month</option>" & vbLf
    s = s & "            <option value=""Day"">Day</option>" & vbLf
    s = s & "          </select>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "        <div class=""input-group"" style=""margin-bottom:15px;"">" & vbLf
    s = s & "          <label class=""input-label"">Apply To (Sync)</label>" & vbLf
    s = s & "          <select id=""newSlicerTarget"" class=""fmt-input""></select>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "      </div>" & vbLf
    s = s & "      <div class=""modal-footer"">" & vbLf
    s = s & "        <button class=""btn-secondary"" onclick=""closeModals()"">Cancel</button>" & vbLf
    s = s & "        <button class=""btn-primary"" style=""background:linear-gradient(135deg, #f43f5e 0%, #e11d48 100%);""" & vbLf
    s = s & "          onclick=""saveNewSlicer()"">Insert Slicer</button>" & vbLf
    s = s & "      </div>" & vbLf
    s = s & "    </div>" & vbLf
    s = s & "  </div>" & vbLf
    s = s & "  <div id=""addTableModal"" class=""modal-overlay"" style=""display: none;"">" & vbLf
    s = s & "    <div class=""modal-box modal-wide"" style=""width: 780px;"">" & vbLf
    s = s & "      <div class=""modal-title"" id=""tblModalTitle"">Smart Data Table Configuration</div>" & vbLf
    s = s & "      <div class=""modal-body"">" & vbLf
    s = s & "" & vbLf
    s = s & "        <div class=""step-head""><span class=""step-badge"">1</span> Basic Details</div>" & vbLf
    s = s & "        <div class=""fmt-row"">" & vbLf
    s = s & "          <div class=""input-group"" style=""flex:1;"">" & vbLf
    s = s & "            <label class=""input-label"">Table Title</label>" & vbLf
    s = s & "            <input type=""text"" id=""tblTitle"" class=""fmt-input"" placeholder=""e.g. Sales by Category"">" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "          <div class=""input-group"" style=""max-width:150px;"">" & vbLf
    s = s & "            <label class=""input-label"">Rows / Page</label>" & vbLf
    s = s & "            <select id=""tblPageSize"" class=""fmt-input"">" & vbLf
    s = s & "              <option value=""5"">5 Rows</option>" & vbLf
    s = s & "              <option value=""15"">15 Rows</option>" & vbLf
    s = s & "              <option value=""28"">28 Rows</option>" & vbLf
    s = s & "              <option value=""29"">29 Rows</option>" & vbLf
    s = s & "              <option value=""30"">30 Rows</option>" & vbLf
    s = s & "              <option value=""31"">31 Rows</option>" & vbLf
    s = s & "              <option value=""50"">50 Rows</option>" & vbLf
    s = s & "            </select>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "" & vbLf
    s = s & "        <div class=""fmt-row size-row"">" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">Widget Width (grid columns)</label>" & vbLf
    s = s & "            <input type=""number"" id=""tblW"" class=""fmt-input"" value=""12"" min=""1"" max=""12"">" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">Widget Height (rows)</label>" & vbLf
    s = s & "            <input type=""number"" id=""tblH"" class=""fmt-input"" value=""5"" min=""1"" max=""12"">" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "" & vbLf
    s = s & "        <div class=""step-head""><span class=""step-badge"">2</span> Table Engine &amp; Logic</div>" & vbLf
    s = s & "        <div class=""pro-tip""><b>&#9432; Pro Tip:</b> Use <b>Aggregated</b> mode to group and summarize data &mdash;" & vbLf
    s = s & "          tick a column, then set its <b>Data Calculation</b> to Group By / Day / Month / Year / Sum / Average / Count." & vbLf
    s = s & "          Columns left at <b>Group By (None)</b> are shown as plain detail. Date columns get Day, Month, Quarter, Year" & vbLf
    s = s & "          and Full Date options; numeric columns get Sum, Average, Count, Min and Max.</div>" & vbLf
    s = s & "" & vbLf
    s = s & "        <div class=""engine-grid"">" & vbLf
    s = s & "          <div class=""engine-card"">" & vbLf
    s = s & "            <div class=""eng-row"">" & vbLf
    s = s & "              <div class=""eng-lab""><b>Data Engine Mode</b><span>Aggregated vs Raw</span></div>" & vbLf
    s = s & "              <select id=""tblMode"" class=""fmt-input"">" & vbLf
    s = s & "                <option value=""agg"">Aggregated</option>" & vbLf
    s = s & "                <option value=""raw"">Raw</option>" & vbLf
    s = s & "              </select>" & vbLf
    s = s & "            </div>" & vbLf
    s = s & "            <div class=""eng-row"">" & vbLf
    s = s & "              <div class=""eng-lab""><b>Show Grand Total</b><span>Bottom metrics row</span></div>" & vbLf
    s = s & "              <label class=""switch""><input type=""checkbox"" id=""tblTotal"" checked><span class=""knob""></span></label>" & vbLf
    s = s & "            </div>" & vbLf
    s = s & "            <div class=""eng-row"">" & vbLf
    s = s & "              <div class=""eng-lab""><b>Wrap Text</b><span>Auto-fit multi-line text</span></div>" & vbLf
    s = s & "              <label class=""switch""><input type=""checkbox"" id=""tblWrap""><span class=""knob""></span></label>" & vbLf
    s = s & "            </div>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "" & vbLf
    s = s & "          <div class=""engine-card"">" & vbLf
    s = s & "            <div class=""eng-lab""><b>&#9733; Top / Bottom Limiter</b><span>Limits table rows based on ranking (Set 0 to show all)</span></div>" & vbLf
    s = s & "            <div class=""lim-row"">" & vbLf
    s = s & "              <select id=""tblLimDir"" class=""fmt-input lim-dir"">" & vbLf
    s = s & "                <option value=""top"">Top</option>" & vbLf
    s = s & "                <option value=""bottom"">Bottom</option>" & vbLf
    s = s & "              </select>" & vbLf
    s = s & "              <input type=""number"" id=""tblLimN"" class=""fmt-input lim-n"" value=""0"" min=""0"">" & vbLf
    s = s & "              <span class=""lim-by"">BY</span>" & vbLf
    s = s & "              <select id=""tblLimCol"" class=""fmt-input lim-col""></select>" & vbLf
    s = s & "            </div>" & vbLf
    s = s & "            <div class=""step-sub"">Date Range Filter (optional)</div>" & vbLf
    s = s & "            <div class=""fmt-row"" style=""margin-bottom:0;"">" & vbLf
    s = s & "              <div class=""input-group"" style=""flex:1;"">" & vbLf
    s = s & "                <label class=""input-label"">Date From</label>" & vbLf
    s = s & "                <input type=""date"" id=""tblFrom"" class=""fmt-input"">" & vbLf
    s = s & "              </div>" & vbLf
    s = s & "              <div class=""input-group"" style=""flex:1;"">" & vbLf
    s = s & "                <label class=""input-label"">Date To</label>" & vbLf
    s = s & "                <input type=""date"" id=""tblTo"" class=""fmt-input"">" & vbLf
    s = s & "              </div>" & vbLf
    s = s & "            </div>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "" & vbLf
    s = s & "        <div class=""step-head""><span class=""step-badge"">3</span> Columns &amp; Data Calculation</div>" & vbLf
    s = s & "        <div class=""col-pick-actions"">" & vbLf
    s = s & "          <button type=""button"" onclick=""tablePickCols(true)"">Select All</button>" & vbLf
    s = s & "          <button type=""button"" onclick=""tablePickCols(false)"">Clear</button>" & vbLf
    s = s & "          <span style=""font-size:11px;color:var(--axis-col);align-self:center;"" id=""tblColCount"">0 selected</span>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "        <div class=""col-rows"" id=""tblCols""></div>" & vbLf
    s = s & "      </div>" & vbLf
    s = s & "      <div class=""modal-footer"">" & vbLf
    s = s & "        <button class=""btn-secondary"" onclick=""closeModals()"">Cancel</button>" & vbLf
    s = s & "        <button id=""tblSubmitBtn"" class=""btn-tbl-generate"" onclick=""saveTable()"">&#9881; Generate Table</button>" & vbLf
    s = s & "      </div>" & vbLf
    s = s & "    </div>" & vbLf
    s = s & "  </div>" & vbLf
    s = s & "  <div id=""updateDataModal"" class=""modal-overlay"" style=""display: none;"">" & vbLf
    s = s & "    <div class=""modal-box"">" & vbLf
    s = s & "      <div class=""modal-title"">Update Dashboard Data</div>" & vbLf
    s = s & "      <div class=""modal-body"">" & vbLf
    s = s & "        <div class=""input-group"" style=""margin-bottom:15px;"">" & vbLf
    s = s & "          <label class=""input-label"">1. Select Excel File</label>" & vbLf
    s = s & "          <input type=""file"" id=""excelFileInput"" class=""fmt-input"" accept="".xlsx, .xls, .csv"">" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "        <div id=""sheetSelectDiv"" class=""input-group"" style=""display:none; margin-bottom:15px;"">" & vbLf
    s = s & "          <label class=""input-label"">2. Select Data Sheet</label>" & vbLf
    s = s & "          <select id=""excelSheetSelect"" class=""fmt-input""></select>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "      </div>" & vbLf
    s = s & "      <div class=""modal-footer"">" & vbLf
    s = s & "        <button class=""btn-secondary"" onclick=""closeModals()"">Cancel</button>" & vbLf
    s = s & "        <button id=""btnProcessExcel"" class=""btn-primary"" style=""display:none;"" onclick=""processSelectedSheet()"">Update" & vbLf
    s = s & "          Dashboard</button>" & vbLf
    s = s & "      </div>" & vbLf
    s = s & "    </div>" & vbLf
    s = s & "  </div>" & vbLf
    s = s & "  <div id=""lblModal"" class=""modal-overlay"" style=""display: none;"">" & vbLf
    s = s & "    <div class=""modal-box"" style=""width: 320px;"">" & vbLf
    s = s & "      <div class=""modal-title"">Label Settings</div>" & vbLf
    s = s & "      <div class=""modal-body"">" & vbLf
    s = s & "        <button id=""btnToggle"" class=""btn-primary"" style=""width:100%; margin-bottom:15px;"" onclick=""toggleLabels()"">Show" & vbLf
    s = s & "          Labels</button>" & vbLf
    s = s & "        <div class=""input-label"" style=""margin-bottom:8px;"">Number Formats</div>" & vbLf
    s = s & "        <button class=""fmt-btn"" onclick=""setLabelMode('raw')"">Actual Value (1,200)</button>" & vbLf
    s = s & "        <button class=""fmt-btn"" onclick=""setLabelMode('auto')"">Auto Format (1K, 1M)</button>" & vbLf
    s = s & "        <button class=""fmt-btn"" onclick=""setLabelMode('K')"">Thousands (50K)</button>" & vbLf
    s = s & "        <button class=""fmt-btn"" onclick=""setLabelMode('M')"">Millions (1.5M)</button>" & vbLf
    s = s & "      </div>" & vbLf
    s = s & "      <div class=""modal-footer"">" & vbLf
    s = s & "        <button class=""btn-secondary"" style=""width:100%;"" onclick=""closeModals()"">Close</button>" & vbLf
    s = s & "      </div>" & vbLf
    s = s & "    </div>" & vbLf
    s = s & "  </div>" & vbLf
    s = s & "  "
    Sec02_BasicModals = s
End Function

'--- Sec03_AddKpiModal : Add KPI modal (title, column, aggregation, icon/image picker)
Private Function Sec03_AddKpiModal() As String
    Dim s As String
    s = s & "<div id=""addKpiModal"" class=""modal-overlay"" style=""display: none;"">" & vbLf
    s = s & "    <div class=""modal-box modal-wide"" style=""width: 500px;"">" & vbLf
    s = s & "      <div class=""modal-title"">Add New KPI</div>" & vbLf
    s = s & "      <div class=""modal-body"">" & vbLf
    s = s & "        <div class=""fmt-row"">" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">KPI Title</label>" & vbLf
    s = s & "            <input type=""text"" id=""newKpiTitle"" class=""fmt-input"" placeholder=""Enter KPI Title"">" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">Select Column</label>" & vbLf
    s = s & "            <select id=""newKpiCol"" class=""fmt-input"">" & vbLf
    s = s & "              <option value=""Order_ID"">Order_ID</option>" & vbLf
    s = s & "              <option value=""Order_Date"">Order_Date</option>" & vbLf
    s = s & "              <option value=""Product"">Product</option>" & vbLf
    s = s & "              <option value=""Category"">Category</option>" & vbLf
    s = s & "              <option value=""Region"">Region</option>" & vbLf
    s = s & "              <option value=""City"">City</option>" & vbLf
    s = s & "              <option value=""State"">State</option>" & vbLf
    s = s & "              <option value=""Customer_Name"">Customer_Name</option>" & vbLf
    s = s & "              <option value=""Order_Type"">Order_Type</option>" & vbLf
    s = s & "              <option value=""Customer_Age"">Customer_Age</option>" & vbLf
    s = s & "              <option value=""Quantity"">Quantity</option>" & vbLf
    s = s & "              <option value=""Price_Per_Unit"">Price_Per_Unit</option>" & vbLf
    s = s & "              <option value=""Total_Amount"">Total_Amount</option>" & vbLf
    s = s & "              <option value=""Return"">Return</option>" & vbLf
    s = s & "            </select>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "        <div class=""fmt-row"">" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">Aggregation</label>" & vbLf
    s = s & "            <select id=""newKpiAgg"" class=""fmt-input"">" & vbLf
    s = s & "              <option value=""sum"">Sum</option>" & vbLf
    s = s & "              <option value=""count"">Count</option>" & vbLf
    s = s & "              <option value=""avg"">Average</option>" & vbLf
    s = s & "            </select>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">Select Icon</label>" & vbLf
    s = s & "            <select id=""newKpiIcon"" class=""fmt-input"">" & vbLf
    s = s & "              <option value="""" selected="""">None</option>" & vbLf
    s = s & "              <option value=""fa-chart-line"">Chart Line</option>" & vbLf
    s = s & "              <option value=""fa-chart-bar"">Bar Chart</option>" & vbLf
    s = s & "              <option value=""fa-chart-pie"">Pie Chart</option>" & vbLf
    s = s & "              <option value=""fa-chart-area"">Area Chart</option>" & vbLf
    s = s & "              <option value=""fa-chart-column"">Column Chart</option>" & vbLf
    s = s & "              <option value=""fa-chart-gantt"">Gantt Chart</option>" & vbLf
    s = s & "              <option value=""fa-arrow-trend-up"">Trend Up</option>" & vbLf
    s = s & "              <option value=""fa-arrow-trend-down"">Trend Down</option>" & vbLf
    s = s & "              <option value=""fa-diagram-project"">Project Diagram</option>" & vbLf
    s = s & "              <option value=""fa-magnifying-glass-chart"">Analyze Chart</option>" & vbLf
    s = s & "              <option value=""fa-indian-rupee-sign"">Rupee (INR)</option>" & vbLf
    s = s & "              <option value=""fa-dollar-sign"">Dollar (USD)</option>" & vbLf
    s = s & "              <option value=""fa-euro-sign"">Euro (EUR)</option>" & vbLf
    s = s & "              <option value=""fa-sterling-sign"">Pound (GBP)</option>" & vbLf
    s = s & "              <option value=""fa-money-bill"">Money</option>" & vbLf
    s = s & "              <option value=""fa-money-bill-wave"">Money Wave</option>" & vbLf
    s = s & "              <option value=""fa-money-bill-trend-up"">Profit/Revenue</option>" & vbLf
    s = s & "              <option value=""fa-coins"">Coins</option>" & vbLf
    s = s & "              <option value=""fa-wallet"">Wallet</option>" & vbLf
    s = s & "              <option value=""fa-credit-card"">Credit Card</option>" & vbLf
    s = s & "              <option value=""fa-receipt"">Receipt</option>" & vbLf
    s = s & "              <option value=""fa-file-invoice"">Invoice</option>" & vbLf
    s = s & "              <option value=""fa-file-invoice-dollar"">Tax/Billing</option>" & vbLf
    s = s & "              <option value=""fa-piggy-bank"">Piggy Bank</option>" & vbLf
    s = s & "              <option value=""fa-vault"">Vault</option>" & vbLf
    s = s & "              <option value=""fa-sack-dollar"">Wealth</option>" & vbLf
    s = s & "              <option value=""fa-scale-balanced"">Scale/Balance</option>" & vbLf
    s = s & "              <option value=""fa-hand-holding-dollar"">Hand &amp; Money</option>" & vbLf
    s = s & "              <option value=""fa-building-columns"">Bank/Institution</option>" & vbLf
    s = s & "              <option value=""fa-store"">Store</option>" & vbLf
    s = s & "              <option value=""fa-shop"">Shop</option>" & vbLf
    s = s & "              <option value=""fa-cart-shopping"">Shopping Cart</option>" & vbLf
    s = s & "              <option value=""fa-cart-plus"">Add to Cart</option>" & vbLf
    s = s & "              <option value=""fa-bag-shopping"">Shopping Bag</option>" & vbLf
    s = s & "              <option value=""fa-basket-shopping"">Basket</option>" & vbLf
    s = s & "              <option value=""fa-truck"">Truck</option>" & vbLf
    s = s & "              <option value=""fa-truck-fast"">Fast Delivery</option>" & vbLf
    s = s & "              <option value=""fa-truck-ramp-box"">Logistics</option>" & vbLf
    s = s & "              <option value=""fa-box"">Box</option>" & vbLf
    s = s & "              <option value=""fa-box-open"">Open Box</option>" & vbLf
    s = s & "              <option value=""fa-boxes-stacked"">Inventory</option>" & vbLf
    s = s & "              <option value=""fa-tags"">Tags</option>" & vbLf
    s = s & "              <option value=""fa-tag"">Tag/Label</option>" & vbLf
    s = s & "              <option value=""fa-barcode"">Barcode</option>" & vbLf
    s = s & "              <option value=""fa-qrcode"">QR Code</option>" & vbLf
    s = s & "              <option value=""fa-gift"">Gift</option>" & vbLf
    s = s & "              <option value=""fa-pallet"">Pallet/Stock</option>" & vbLf
    s = s & "              <option value=""fa-warehouse"">Warehouse</option>" & vbLf
    s = s & "              <option value=""fa-plane"">Airplane</option>" & vbLf
    s = s & "              <option value=""fa-ship"">Ship</option>" & vbLf
    s = s & "              <option value=""fa-user"">User</option>" & vbLf
    s = s & "              <option value=""fa-users"">Users/Team</option>" & vbLf
    s = s & "              <option value=""fa-user-group"">User Group</option>" & vbLf
    s = s & "              <option value=""fa-user-tie"">Executive</option>" & vbLf
    s = s & "              <option value=""fa-user-check"">Verified User</option>" & vbLf
    s = s & "              <option value=""fa-user-plus"">Add User</option>" & vbLf
    s = s & "              <option value=""fa-user-minus"">Remove User</option>" & vbLf
    s = s & "              <option value=""fa-user-gear"">Admin User</option>" & vbLf
    s = s & "              <option value=""fa-user-shield"">Secure User</option>" & vbLf
    s = s & "              <option value=""fa-id-badge"">ID Badge</option>" & vbLf
    s = s & "              <option value=""fa-id-card"">ID Card</option>" & vbLf
    s = s & "              <option value=""fa-address-book"">Address Book</option>" & vbLf
    s = s & "              <option value=""fa-building"">Building/Company</option>" & vbLf
    s = s & "              <option value=""fa-city"">City</option>" & vbLf
    s = s & "              <option value=""fa-briefcase"">Briefcase</option>" & vbLf
    s = s & "              <option value=""fa-envelope"">Email</option>" & vbLf
    s = s & "              <option value=""fa-envelope-open-text"">Read Email</option>" & vbLf
    s = s & "              <option value=""fa-phone"">Phone</option>" & vbLf
    s = s & "              <option value=""fa-phone-volume"">Active Call</option>" & vbLf
    s = s & "              <option value=""fa-comment"">Chat/Comment</option>" & vbLf
    s = s & "              <option value=""fa-comments"">Conversations</option>" & vbLf
    s = s & "              <option value=""fa-paper-plane"">Send/Launch</option>" & vbLf
    s = s & "              <option value=""fa-clipboard"">Clipboard</option>" & vbLf
    s = s & "              <option value=""fa-clipboard-list"">Task List</option>" & vbLf
    s = s & "              <option value=""fa-clipboard-check"">Tasks Done</option>" & vbLf
    s = s & "              <option value=""fa-list-check"">Checklist</option>" & vbLf
    s = s & "              <option value=""fa-list-ol"">Numbered List</option>" & vbLf
    s = s & "              <option value=""fa-list-ul"">Bullet List</option>" & vbLf
    s = s & "              <option value=""fa-table"">Data Table</option>" & vbLf
    s = s & "              <option value=""fa-table-list"">List Table</option>" & vbLf
    s = s & "              <option value=""fa-calendar"">Calendar</option>" & vbLf
    s = s & "              <option value=""fa-calendar-days"">Calendar Days</option>" & vbLf
    s = s & "              <option value=""fa-calendar-check"">Event Done</option>" & vbLf
    s = s & "              <option value=""fa-clock"">Clock</option>" & vbLf
    s = s & "              <option value=""fa-stopwatch"">Stopwatch</option>" & vbLf
    s = s & "              <option value=""fa-hourglass-half"">Hourglass</option>" & vbLf
    s = s & "              <option value=""fa-business-time"">Business Hours</option>" & vbLf
    s = s & "              <option value=""fa-timeline"">Timeline</option>" & vbLf
    s = s & "              <option value=""fa-bullseye"">Target/Goal</option>" & vbLf
    s = s & "              <option value=""fa-trophy"">Trophy</option>" & vbLf
    s = s & "              <option value=""fa-handshake"">Handshake</option>" & vbLf
    s = s & "              <option value=""fa-database"">Database</option>" & vbLf
    s = s & "              <option value=""fa-server"">Server</option>" & vbLf
    s = s & "              <option value=""fa-cloud"">Cloud</option>" & vbLf
    s = s & "              <option value=""fa-cloud-arrow-up"">Cloud Upload</option>" & vbLf
    s = s & "              <option value=""fa-cloud-arrow-down"">Cloud Download</option>" & vbLf
    s = s & "              <option value=""fa-network-wired"">Network</option>" & vbLf
    s = s & "              <option value=""fa-laptop"">Laptop</option>" & vbLf
    s = s & "              <option value=""fa-desktop"">Desktop</option>" & vbLf
    s = s & "              <option value=""fa-mobile-screen"">Mobile</option>" & vbLf
    s = s & "              <option value=""fa-tablet-screen"">Tablet</option>" & vbLf
    s = s & "              <option value=""fa-microchip"">Processor</option>" & vbLf
    s = s & "              <option value=""fa-memory"">Memory/RAM</option>" & vbLf
    s = s & "              <option value=""fa-code"">Code</option>" & vbLf
    s = s & "              <option value=""fa-code-branch"">Version Control</option>" & vbLf
    s = s & "              <option value=""fa-sitemap"">Sitemap/Flow</option>" & vbLf
    s = s & "              <option value=""fa-wifi"">Wi-Fi</option>" & vbLf
    s = s & "              <option value=""fa-globe"">Globe</option>" & vbLf
    s = s & "              <option value=""fa-earth-americas"">Earth</option>" & vbLf
    s = s & "              <option value=""fa-location-dot"">Location</option>" & vbLf
    s = s & "              <option value=""fa-map-location-dot"">Map</option>" & vbLf
    s = s & "              <option value=""fa-map-pin"">Map Pin</option>" & vbLf
    s = s & "              <option value=""fa-compass"">Compass</option>" & vbLf
    s = s & "              <option value=""fa-satellite-dish"">Satellite</option>" & vbLf
    s = s & "              <option value=""fa-star"">Star</option>" & vbLf
    s = s & "              <option value=""fa-star-half-stroke"">Half Star</option>" & vbLf
    s = s & "              <option value=""fa-heart"">Heart</option>" & vbLf
    s = s & "              <option value=""fa-heart-pulse"">Heart Rate</option>" & vbLf
    s = s & "              <option value=""fa-fire"">Fire/Hot</option>" & vbLf
    s = s & "              <option value=""fa-bolt"">Flash/Speed</option>" & vbLf
    s = s & "              <option value=""fa-thumbs-up"">Thumbs Up</option>" & vbLf
    s = s & "              <option value=""fa-thumbs-down"">Thumbs Down</option>" & vbLf
    s = s & "              <option value=""fa-crown"">Crown/Premium</option>" & vbLf
    s = s & "              <option value=""fa-medal"">Medal</option>" & vbLf
    s = s & "              <option value=""fa-award"">Award</option>" & vbLf
    s = s & "              <option value=""fa-ribbon"">Ribbon</option>" & vbLf
    s = s & "              <option value=""fa-ranking-star"">Ranking</option>" & vbLf
    s = s & "              <option value=""fa-percent"">Percentage (%)</option>" & vbLf
    s = s & "              <option value=""fa-calculator"">Calculator</option>" & vbLf
    s = s & "              <option value=""fa-scale-unbalanced"">Unbalanced</option>" & vbLf
    s = s & "              <option value=""fa-circle-half-stroke"">Completion/Pie</option>" & vbLf
    s = s & "              <option value=""fa-shield"">Shield</option>" & vbLf
    s = s & "              <option value=""fa-shield-halved"">Security</option>" & vbLf
    s = s & "              <option value=""fa-lock"">Lock</option>" & vbLf
    s = s & "              <option value=""fa-unlock"">Unlock</option>" & vbLf
    s = s & "              <option value=""fa-key"">Key</option>" & vbLf
    s = s & "              <option value=""fa-gear"">Setting</option>" & vbLf
    s = s & "              <option value=""fa-gears"">System/Gears</option>" & vbLf
    s = s & "              <option value=""fa-sliders"">Adjust/Sliders</option>" & vbLf
    s = s & "              <option value=""fa-wrench"">Wrench</option>" & vbLf
    s = s & "              <option value=""fa-screwdriver-wrench"">Repair/Tools</option>" & vbLf
    s = s & "              <option value=""fa-triangle-exclamation"">Warning</option>" & vbLf
    s = s & "              <option value=""fa-circle-exclamation"">Alert</option>" & vbLf
    s = s & "              <option value=""fa-circle-check"">Success</option>" & vbLf
    s = s & "              <option value=""fa-circle-info"">Information</option>" & vbLf
    s = s & "              <option value=""fa-circle-xmark"">Error/Cancel</option>" & vbLf
    s = s & "              <option value=""fa-file"">File</option>" & vbLf
    s = s & "              <option value=""fa-file-lines"">Document</option>" & vbLf
    s = s & "              <option value=""fa-file-csv"">CSV File</option>" & vbLf
    s = s & "              <option value=""fa-file-excel"">Excel File</option>" & vbLf
    s = s & "              <option value=""fa-file-pdf"">PDF File</option>" & vbLf
    s = s & "              <option value=""fa-folder"">Folder</option>" & vbLf
    s = s & "              <option value=""fa-folder-open"">Open Folder</option>" & vbLf
    s = s & "              <option value=""fa-folder-tree"">Directory</option>" & vbLf
    s = s & "              <option value=""fa-copy"">Copy</option>" & vbLf
    s = s & "              <option value=""fa-paste"">Paste</option>" & vbLf
    s = s & "              <option value=""fa-download"">Download</option>" & vbLf
    s = s & "              <option value=""fa-upload"">Upload</option>" & vbLf
    s = s & "              <option value=""fa-print"">Print</option>" & vbLf
    s = s & "              <option value=""fa-trash"">Delete/Trash</option>" & vbLf
    s = s & "              <option value=""fa-pen"">Pen</option>" & vbLf
    s = s & "              <option value=""fa-pen-to-square"">Edit</option>" & vbLf
    s = s & "              <option value=""fa-eye"">View/Visibility</option>" & vbLf
    s = s & "              <option value=""fa-eye-slash"">Hide</option>" & vbLf
    s = s & "              <option value=""fa-magnifying-glass"">Search</option>" & vbLf
    s = s & "              <option value=""fa-filter"">Filter</option>" & vbLf
    s = s & "              <option value=""fa-sort"">Sort</option>" & vbLf
    s = s & "              <option value=""fa-share-nodes"">Share</option>" & vbLf
    s = s & "              <option value=""fa-arrows-rotate"">Refresh/Sync</option>" & vbLf
    s = s & "              <option value=""fa-lightbulb"">Idea</option>" & vbLf
    s = s & "              <option value=""fa-leaf"">Eco/Nature</option>" & vbLf
    s = s & "              <option value=""fa-droplet"">Water/Liquid</option>" & vbLf
    s = s & "              <option value=""fa-sun"">Sun/Day</option>" & vbLf
    s = s & "              <option value=""fa-moon"">Moon/Night</option>" & vbLf
    s = s & "              <option value=""fa-snowflake"">Cold/Snow</option>" & vbLf
    s = s & "              <option value=""fa-car"">Car</option>" & vbLf
    s = s & "              <option value=""fa-bus"">Bus</option>" & vbLf
    s = s & "              <option value=""fa-train"">Train</option>" & vbLf
    s = s & "              <option value=""fa-bicycle"">Bicycle</option>" & vbLf
    s = s & "              <option value=""fa-motorcycle"">Motorcycle</option>" & vbLf
    s = s & "              <option value=""fa-hospital"">Hospital</option>" & vbLf
    s = s & "              <option value=""fa-stethoscope"">Health</option>" & vbLf
    s = s & "              <option value=""fa-pills"">Medicine</option>" & vbLf
    s = s & "              <option value=""fa-graduation-cap"">Education</option>" & vbLf
    s = s & "              <option value=""fa-book"">Book</option>" & vbLf
    s = s & "              <option value=""fa-laptop-code"">E-learning</option>" & vbLf
    s = s & "              <option value=""fa-mug-hot"">Coffee/Break</option>" & vbLf
    s = s & "              <option value=""fa-utensils"">Food/Dining</option>" & vbLf
    s = s & "              <option value=""fa-wine-glass"">Drinks</option>" & vbLf
    s = s & "              <option value=""fa-battery-full"">Battery Full</option>" & vbLf
    s = s & "              <option value=""fa-battery-half"">Battery Half</option>" & vbLf
    s = s & "              <option value=""fa-battery-empty"">Battery Empty</option>" & vbLf
    s = s & "              <option value=""fa-temperature-half"">Temperature</option>" & vbLf
    s = s & "              <option value=""fa-link"">Link/URL</option>" & vbLf
    s = s & "              <option value=""fa-link-slash"">Unlink</option>" & vbLf
    s = s & "              <option value=""fa-paperclip"">Attachment</option>" & vbLf
    s = s & "              <option value=""fa-bookmark"">Bookmark</option>" & vbLf
    s = s & "              <option value=""fa-bell"">Notification</option>" & vbLf
    s = s & "              <option value=""fa-microphone"">Microphone</option>" & vbLf
    s = s & "              <option value=""fa-volume-high"">Volume Up</option>" & vbLf
    s = s & "              <option value=""fa-volume-xmark"">Mute</option>" & vbLf
    s = s & "              <option value=""fa-camera"">Camera</option>" & vbLf
    s = s & "              <option value=""fa-video"">Video</option>" & vbLf
    s = s & "              <option value=""fa-image"">Image</option>" & vbLf
    s = s & "              <option value=""fa-check"">Check</option>" & vbLf
    s = s & "              <option value=""fa-xmark"">Cross/Close</option>" & vbLf
    s = s & "              <option value=""fa-plus"">Add/Plus</option>" & vbLf
    s = s & "              <option value=""fa-minus"">Subtract</option>" & vbLf
    s = s & "              <option value=""fa-arrow-right"">Arrow Right</option>" & vbLf
    s = s & "              <option value=""fa-arrow-left"">Arrow Left</option>" & vbLf
    s = s & "              <option value=""fa-arrow-up"">Arrow Up</option>" & vbLf
    s = s & "              <option value=""fa-arrow-down"">Arrow Down</option>" & vbLf
    s = s & "              <option value=""fa-caret-up"">Caret Up</option>" & vbLf
    s = s & "              <option value=""fa-caret-down"">Caret Down</option>" & vbLf
    s = s & "            </select>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "        <div class=""input-group"" style=""margin-bottom:15px;"">" & vbLf
    s = s & "          <label class=""input-label"">OR Upload Image from PC</label>" & vbLf
    s = s & "          <input type=""file"" id=""newKpiImage"" class=""fmt-input"" accept=""image/*"" onchange=""handleKpiImageUpload(event)"">" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "      </div>" & vbLf
    s = s & "      <div class=""modal-footer"">" & vbLf
    s = s & "        <button class=""btn-secondary"" onclick=""closeModals()"">Cancel</button>" & vbLf
    s = s & "        <button class=""btn-primary"" style=""background:linear-gradient(135deg, #4f46e5 0%, #4338ca 100%);""" & vbLf
    s = s & "          onclick=""saveNewKPI()"">Add KPI</button>" & vbLf
    s = s & "      </div>" & vbLf
    s = s & "    </div>" & vbLf
    s = s & "  </div>" & vbLf
    s = s & "  "
    Sec03_AddKpiModal = s
End Function

'--- Sec04_EditKpiModal : Edit KPI modal
Private Function Sec04_EditKpiModal() As String
    Dim s As String
    s = s & "<div id=""editKpiModal"" class=""modal-overlay"" style=""display: none;"">" & vbLf
    s = s & "    <div class=""modal-box modal-wide"" style=""width: 500px;"">" & vbLf
    s = s & "      <div class=""modal-title"">Edit KPI Settings</div>" & vbLf
    s = s & "      <div class=""modal-body"">" & vbLf
    s = s & "        <div class=""fmt-row"">" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">KPI Title</label>" & vbLf
    s = s & "            <input type=""text"" id=""editKpiTitle"" class=""fmt-input"" placeholder=""Enter KPI Title"">" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">Select Column</label>" & vbLf
    s = s & "            <select id=""editKpiCol"" class=""fmt-input"">" & vbLf
    s = s & "              <option value=""Order_ID"">Order_ID</option>" & vbLf
    s = s & "              <option value=""Order_Date"">Order_Date</option>" & vbLf
    s = s & "              <option value=""Product"">Product</option>" & vbLf
    s = s & "              <option value=""Category"">Category</option>" & vbLf
    s = s & "              <option value=""Region"">Region</option>" & vbLf
    s = s & "              <option value=""City"">City</option>" & vbLf
    s = s & "              <option value=""State"">State</option>" & vbLf
    s = s & "              <option value=""Customer_Name"">Customer_Name</option>" & vbLf
    s = s & "              <option value=""Order_Type"">Order_Type</option>" & vbLf
    s = s & "              <option value=""Customer_Age"">Customer_Age</option>" & vbLf
    s = s & "              <option value=""Quantity"">Quantity</option>" & vbLf
    s = s & "              <option value=""Price_Per_Unit"">Price_Per_Unit</option>" & vbLf
    s = s & "              <option value=""Total_Amount"">Total_Amount</option>" & vbLf
    s = s & "              <option value=""Return"">Return</option>" & vbLf
    s = s & "            </select>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "        <div class=""fmt-row"">" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">Aggregation</label>" & vbLf
    s = s & "            <select id=""editKpiAgg"" class=""fmt-input"">" & vbLf
    s = s & "              <option value=""sum"">Sum</option>" & vbLf
    s = s & "              <option value=""count"">Count</option>" & vbLf
    s = s & "              <option value=""avg"">Average</option>" & vbLf
    s = s & "            </select>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">Select Icon</label>" & vbLf
    s = s & "            <select id=""editKpiIcon"" class=""fmt-input"">" & vbLf
    s = s & "              <option value="""">None</option>" & vbLf
    s = s & "              <option value=""fa-chart-line"">Chart Line</option>" & vbLf
    s = s & "              <option value=""fa-chart-bar"">Bar Chart</option>" & vbLf
    s = s & "              <option value=""fa-chart-pie"">Pie Chart</option>" & vbLf
    s = s & "              <option value=""fa-chart-area"">Area Chart</option>" & vbLf
    s = s & "              <option value=""fa-chart-column"">Column Chart</option>" & vbLf
    s = s & "              <option value=""fa-chart-gantt"">Gantt Chart</option>" & vbLf
    s = s & "              <option value=""fa-arrow-trend-up"">Trend Up</option>" & vbLf
    s = s & "              <option value=""fa-arrow-trend-down"">Trend Down</option>" & vbLf
    s = s & "              <option value=""fa-diagram-project"">Project Diagram</option>" & vbLf
    s = s & "              <option value=""fa-magnifying-glass-chart"">Analyze Chart</option>" & vbLf
    s = s & "              <option value=""fa-indian-rupee-sign"">Rupee (INR)</option>" & vbLf
    s = s & "              <option value=""fa-dollar-sign"">Dollar (USD)</option>" & vbLf
    s = s & "              <option value=""fa-euro-sign"">Euro (EUR)</option>" & vbLf
    s = s & "              <option value=""fa-sterling-sign"">Pound (GBP)</option>" & vbLf
    s = s & "              <option value=""fa-money-bill"">Money</option>" & vbLf
    s = s & "              <option value=""fa-money-bill-wave"">Money Wave</option>" & vbLf
    s = s & "              <option value=""fa-money-bill-trend-up"">Profit/Revenue</option>" & vbLf
    s = s & "              <option value=""fa-coins"">Coins</option>" & vbLf
    s = s & "              <option value=""fa-wallet"">Wallet</option>" & vbLf
    s = s & "              <option value=""fa-credit-card"">Credit Card</option>" & vbLf
    s = s & "              <option value=""fa-receipt"">Receipt</option>" & vbLf
    s = s & "              <option value=""fa-file-invoice"">Invoice</option>" & vbLf
    s = s & "              <option value=""fa-file-invoice-dollar"">Tax/Billing</option>" & vbLf
    s = s & "              <option value=""fa-piggy-bank"">Piggy Bank</option>" & vbLf
    s = s & "              <option value=""fa-vault"">Vault</option>" & vbLf
    s = s & "              <option value=""fa-sack-dollar"">Wealth</option>" & vbLf
    s = s & "              <option value=""fa-scale-balanced"">Scale/Balance</option>" & vbLf
    s = s & "              <option value=""fa-hand-holding-dollar"">Hand &amp; Money</option>" & vbLf
    s = s & "              <option value=""fa-building-columns"">Bank/Institution</option>" & vbLf
    s = s & "              <option value=""fa-store"">Store</option>" & vbLf
    s = s & "              <option value=""fa-shop"">Shop</option>" & vbLf
    s = s & "              <option value=""fa-cart-shopping"">Shopping Cart</option>" & vbLf
    s = s & "              <option value=""fa-cart-plus"" selected="""">Add to Cart</option>" & vbLf
    s = s & "              <option value=""fa-bag-shopping"">Shopping Bag</option>" & vbLf
    s = s & "              <option value=""fa-basket-shopping"">Basket</option>" & vbLf
    s = s & "              <option value=""fa-truck"">Truck</option>" & vbLf
    s = s & "              <option value=""fa-truck-fast"">Fast Delivery</option>" & vbLf
    s = s & "              <option value=""fa-truck-ramp-box"">Logistics</option>" & vbLf
    s = s & "              <option value=""fa-box"">Box</option>" & vbLf
    s = s & "              <option value=""fa-box-open"">Open Box</option>" & vbLf
    s = s & "              <option value=""fa-boxes-stacked"">Inventory</option>" & vbLf
    s = s & "              <option value=""fa-tags"">Tags</option>" & vbLf
    s = s & "              <option value=""fa-tag"">Tag/Label</option>" & vbLf
    s = s & "              <option value=""fa-barcode"">Barcode</option>" & vbLf
    s = s & "              <option value=""fa-qrcode"">QR Code</option>" & vbLf
    s = s & "              <option value=""fa-gift"">Gift</option>" & vbLf
    s = s & "              <option value=""fa-pallet"">Pallet/Stock</option>" & vbLf
    s = s & "              <option value=""fa-warehouse"">Warehouse</option>" & vbLf
    s = s & "              <option value=""fa-plane"">Airplane</option>" & vbLf
    s = s & "              <option value=""fa-ship"">Ship</option>" & vbLf
    s = s & "              <option value=""fa-user"">User</option>" & vbLf
    s = s & "              <option value=""fa-users"">Users/Team</option>" & vbLf
    s = s & "              <option value=""fa-user-group"">User Group</option>" & vbLf
    s = s & "              <option value=""fa-user-tie"">Executive</option>" & vbLf
    s = s & "              <option value=""fa-user-check"">Verified User</option>" & vbLf
    s = s & "              <option value=""fa-user-plus"">Add User</option>" & vbLf
    s = s & "              <option value=""fa-user-minus"">Remove User</option>" & vbLf
    s = s & "              <option value=""fa-user-gear"">Admin User</option>" & vbLf
    s = s & "              <option value=""fa-user-shield"">Secure User</option>" & vbLf
    s = s & "              <option value=""fa-id-badge"">ID Badge</option>" & vbLf
    s = s & "              <option value=""fa-id-card"">ID Card</option>" & vbLf
    s = s & "              <option value=""fa-address-book"">Address Book</option>" & vbLf
    s = s & "              <option value=""fa-building"">Building/Company</option>" & vbLf
    s = s & "              <option value=""fa-city"">City</option>" & vbLf
    s = s & "              <option value=""fa-briefcase"">Briefcase</option>" & vbLf
    s = s & "              <option value=""fa-envelope"">Email</option>" & vbLf
    s = s & "              <option value=""fa-envelope-open-text"">Read Email</option>" & vbLf
    s = s & "              <option value=""fa-phone"">Phone</option>" & vbLf
    s = s & "              <option value=""fa-phone-volume"">Active Call</option>" & vbLf
    s = s & "              <option value=""fa-comment"">Chat/Comment</option>" & vbLf
    s = s & "              <option value=""fa-comments"">Conversations</option>" & vbLf
    s = s & "              <option value=""fa-paper-plane"">Send/Launch</option>" & vbLf
    s = s & "              <option value=""fa-clipboard"">Clipboard</option>" & vbLf
    s = s & "              <option value=""fa-clipboard-list"">Task List</option>" & vbLf
    s = s & "              <option value=""fa-clipboard-check"">Tasks Done</option>" & vbLf
    s = s & "              <option value=""fa-list-check"">Checklist</option>" & vbLf
    s = s & "              <option value=""fa-list-ol"">Numbered List</option>" & vbLf
    s = s & "              <option value=""fa-list-ul"">Bullet List</option>" & vbLf
    s = s & "              <option value=""fa-table"">Data Table</option>" & vbLf
    s = s & "              <option value=""fa-table-list"">List Table</option>" & vbLf
    s = s & "              <option value=""fa-calendar"">Calendar</option>" & vbLf
    s = s & "              <option value=""fa-calendar-days"">Calendar Days</option>" & vbLf
    s = s & "              <option value=""fa-calendar-check"">Event Done</option>" & vbLf
    s = s & "              <option value=""fa-clock"">Clock</option>" & vbLf
    s = s & "              <option value=""fa-stopwatch"">Stopwatch</option>" & vbLf
    s = s & "              <option value=""fa-hourglass-half"">Hourglass</option>" & vbLf
    s = s & "              <option value=""fa-business-time"">Business Hours</option>" & vbLf
    s = s & "              <option value=""fa-timeline"">Timeline</option>" & vbLf
    s = s & "              <option value=""fa-bullseye"">Target/Goal</option>" & vbLf
    s = s & "              <option value=""fa-trophy"">Trophy</option>" & vbLf
    s = s & "              <option value=""fa-handshake"">Handshake</option>" & vbLf
    s = s & "              <option value=""fa-database"">Database</option>" & vbLf
    s = s & "              <option value=""fa-server"">Server</option>" & vbLf
    s = s & "              <option value=""fa-cloud"">Cloud</option>" & vbLf
    s = s & "              <option value=""fa-cloud-arrow-up"">Cloud Upload</option>" & vbLf
    s = s & "              <option value=""fa-cloud-arrow-down"">Cloud Download</option>" & vbLf
    s = s & "              <option value=""fa-network-wired"">Network</option>" & vbLf
    s = s & "              <option value=""fa-laptop"">Laptop</option>" & vbLf
    s = s & "              <option value=""fa-desktop"">Desktop</option>" & vbLf
    s = s & "              <option value=""fa-mobile-screen"">Mobile</option>" & vbLf
    s = s & "              <option value=""fa-tablet-screen"">Tablet</option>" & vbLf
    s = s & "              <option value=""fa-microchip"">Processor</option>" & vbLf
    s = s & "              <option value=""fa-memory"">Memory/RAM</option>" & vbLf
    s = s & "              <option value=""fa-code"">Code</option>" & vbLf
    s = s & "              <option value=""fa-code-branch"">Version Control</option>" & vbLf
    s = s & "              <option value=""fa-sitemap"">Sitemap/Flow</option>" & vbLf
    s = s & "              <option value=""fa-wifi"">Wi-Fi</option>" & vbLf
    s = s & "              <option value=""fa-globe"">Globe</option>" & vbLf
    s = s & "              <option value=""fa-earth-americas"">Earth</option>" & vbLf
    s = s & "              <option value=""fa-location-dot"">Location</option>" & vbLf
    s = s & "              <option value=""fa-map-location-dot"">Map</option>" & vbLf
    s = s & "              <option value=""fa-map-pin"">Map Pin</option>" & vbLf
    s = s & "              <option value=""fa-compass"">Compass</option>" & vbLf
    s = s & "              <option value=""fa-satellite-dish"">Satellite</option>" & vbLf
    s = s & "              <option value=""fa-star"">Star</option>" & vbLf
    s = s & "              <option value=""fa-star-half-stroke"">Half Star</option>" & vbLf
    s = s & "              <option value=""fa-heart"">Heart</option>" & vbLf
    s = s & "              <option value=""fa-heart-pulse"">Heart Rate</option>" & vbLf
    s = s & "              <option value=""fa-fire"">Fire/Hot</option>" & vbLf
    s = s & "              <option value=""fa-bolt"">Flash/Speed</option>" & vbLf
    s = s & "              <option value=""fa-thumbs-up"">Thumbs Up</option>" & vbLf
    s = s & "              <option value=""fa-thumbs-down"">Thumbs Down</option>" & vbLf
    s = s & "              <option value=""fa-crown"">Crown/Premium</option>" & vbLf
    s = s & "              <option value=""fa-medal"">Medal</option>" & vbLf
    s = s & "              <option value=""fa-award"">Award</option>" & vbLf
    s = s & "              <option value=""fa-ribbon"">Ribbon</option>" & vbLf
    s = s & "              <option value=""fa-ranking-star"">Ranking</option>" & vbLf
    s = s & "              <option value=""fa-percent"">Percentage (%)</option>" & vbLf
    s = s & "              <option value=""fa-calculator"">Calculator</option>" & vbLf
    s = s & "              <option value=""fa-scale-unbalanced"">Unbalanced</option>" & vbLf
    s = s & "              <option value=""fa-circle-half-stroke"">Completion/Pie</option>" & vbLf
    s = s & "              <option value=""fa-shield"">Shield</option>" & vbLf
    s = s & "              <option value=""fa-shield-halved"">Security</option>" & vbLf
    s = s & "              <option value=""fa-lock"">Lock</option>" & vbLf
    s = s & "              <option value=""fa-unlock"">Unlock</option>" & vbLf
    s = s & "              <option value=""fa-key"">Key</option>" & vbLf
    s = s & "              <option value=""fa-gear"">Setting</option>" & vbLf
    s = s & "              <option value=""fa-gears"">System/Gears</option>" & vbLf
    s = s & "              <option value=""fa-sliders"">Adjust/Sliders</option>" & vbLf
    s = s & "              <option value=""fa-wrench"">Wrench</option>" & vbLf
    s = s & "              <option value=""fa-screwdriver-wrench"">Repair/Tools</option>" & vbLf
    s = s & "              <option value=""fa-triangle-exclamation"">Warning</option>" & vbLf
    s = s & "              <option value=""fa-circle-exclamation"">Alert</option>" & vbLf
    s = s & "              <option value=""fa-circle-check"">Success</option>" & vbLf
    s = s & "              <option value=""fa-circle-info"">Information</option>" & vbLf
    s = s & "              <option value=""fa-circle-xmark"">Error/Cancel</option>" & vbLf
    s = s & "              <option value=""fa-file"">File</option>" & vbLf
    s = s & "              <option value=""fa-file-lines"">Document</option>" & vbLf
    s = s & "              <option value=""fa-file-csv"">CSV File</option>" & vbLf
    s = s & "              <option value=""fa-file-excel"">Excel File</option>" & vbLf
    s = s & "              <option value=""fa-file-pdf"">PDF File</option>" & vbLf
    s = s & "              <option value=""fa-folder"">Folder</option>" & vbLf
    s = s & "              <option value=""fa-folder-open"">Open Folder</option>" & vbLf
    s = s & "              <option value=""fa-folder-tree"">Directory</option>" & vbLf
    s = s & "              <option value=""fa-copy"">Copy</option>" & vbLf
    s = s & "              <option value=""fa-paste"">Paste</option>" & vbLf
    s = s & "              <option value=""fa-download"">Download</option>" & vbLf
    s = s & "              <option value=""fa-upload"">Upload</option>" & vbLf
    s = s & "              <option value=""fa-print"">Print</option>" & vbLf
    s = s & "              <option value=""fa-trash"">Delete/Trash</option>" & vbLf
    s = s & "              <option value=""fa-pen"">Pen</option>" & vbLf
    s = s & "              <option value=""fa-pen-to-square"">Edit</option>" & vbLf
    s = s & "              <option value=""fa-eye"">View/Visibility</option>" & vbLf
    s = s & "              <option value=""fa-eye-slash"">Hide</option>" & vbLf
    s = s & "              <option value=""fa-magnifying-glass"">Search</option>" & vbLf
    s = s & "              <option value=""fa-filter"">Filter</option>" & vbLf
    s = s & "              <option value=""fa-sort"">Sort</option>" & vbLf
    s = s & "              <option value=""fa-share-nodes"">Share</option>" & vbLf
    s = s & "              <option value=""fa-arrows-rotate"">Refresh/Sync</option>" & vbLf
    s = s & "              <option value=""fa-lightbulb"">Idea</option>" & vbLf
    s = s & "              <option value=""fa-leaf"">Eco/Nature</option>" & vbLf
    s = s & "              <option value=""fa-droplet"">Water/Liquid</option>" & vbLf
    s = s & "              <option value=""fa-sun"">Sun/Day</option>" & vbLf
    s = s & "              <option value=""fa-moon"">Moon/Night</option>" & vbLf
    s = s & "              <option value=""fa-snowflake"">Cold/Snow</option>" & vbLf
    s = s & "              <option value=""fa-car"">Car</option>" & vbLf
    s = s & "              <option value=""fa-bus"">Bus</option>" & vbLf
    s = s & "              <option value=""fa-train"">Train</option>" & vbLf
    s = s & "              <option value=""fa-bicycle"">Bicycle</option>" & vbLf
    s = s & "              <option value=""fa-motorcycle"">Motorcycle</option>" & vbLf
    s = s & "              <option value=""fa-hospital"">Hospital</option>" & vbLf
    s = s & "              <option value=""fa-stethoscope"">Health</option>" & vbLf
    s = s & "              <option value=""fa-pills"">Medicine</option>" & vbLf
    s = s & "              <option value=""fa-graduation-cap"">Education</option>" & vbLf
    s = s & "              <option value=""fa-book"">Book</option>" & vbLf
    s = s & "              <option value=""fa-laptop-code"">E-learning</option>" & vbLf
    s = s & "              <option value=""fa-mug-hot"">Coffee/Break</option>" & vbLf
    s = s & "              <option value=""fa-utensils"">Food/Dining</option>" & vbLf
    s = s & "              <option value=""fa-wine-glass"">Drinks</option>" & vbLf
    s = s & "              <option value=""fa-battery-full"">Battery Full</option>" & vbLf
    s = s & "              <option value=""fa-battery-half"">Battery Half</option>" & vbLf
    s = s & "              <option value=""fa-battery-empty"">Battery Empty</option>" & vbLf
    s = s & "              <option value=""fa-temperature-half"">Temperature</option>" & vbLf
    s = s & "              <option value=""fa-link"">Link/URL</option>" & vbLf
    s = s & "              <option value=""fa-link-slash"">Unlink</option>" & vbLf
    s = s & "              <option value=""fa-paperclip"">Attachment</option>" & vbLf
    s = s & "              <option value=""fa-bookmark"">Bookmark</option>" & vbLf
    s = s & "              <option value=""fa-bell"">Notification</option>" & vbLf
    s = s & "              <option value=""fa-microphone"">Microphone</option>" & vbLf
    s = s & "              <option value=""fa-volume-high"">Volume Up</option>" & vbLf
    s = s & "              <option value=""fa-volume-xmark"">Mute</option>" & vbLf
    s = s & "              <option value=""fa-camera"">Camera</option>" & vbLf
    s = s & "              <option value=""fa-video"">Video</option>" & vbLf
    s = s & "              <option value=""fa-image"">Image</option>" & vbLf
    s = s & "              <option value=""fa-check"">Check</option>" & vbLf
    s = s & "              <option value=""fa-xmark"">Cross/Close</option>" & vbLf
    s = s & "              <option value=""fa-plus"">Add/Plus</option>" & vbLf
    s = s & "              <option value=""fa-minus"">Subtract</option>" & vbLf
    s = s & "              <option value=""fa-arrow-right"">Arrow Right</option>" & vbLf
    s = s & "              <option value=""fa-arrow-left"">Arrow Left</option>" & vbLf
    s = s & "              <option value=""fa-arrow-up"">Arrow Up</option>" & vbLf
    s = s & "              <option value=""fa-arrow-down"">Arrow Down</option>" & vbLf
    s = s & "              <option value=""fa-caret-up"">Caret Up</option>" & vbLf
    s = s & "              <option value=""fa-caret-down"">Caret Down</option>" & vbLf
    s = s & "            </select>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "        <div class=""input-group"" style=""margin-bottom:15px;"">" & vbLf
    s = s & "          <label class=""input-label"">Replace Image (Upload new)</label>" & vbLf
    s = s & "          <input type=""file"" id=""editKpiImage"" class=""fmt-input"" accept=""image/*""" & vbLf
    s = s & "            onchange=""handleKpiImageUpload(event)"">" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "      </div>" & vbLf
    s = s & "      <div class=""modal-footer"">" & vbLf
    s = s & "        <button class=""btn-secondary"" onclick=""closeModals()"">Cancel</button>" & vbLf
    s = s & "        <button class=""btn-primary"" style=""background:linear-gradient(135deg, #0ea5e9 0%, #2563eb 100%);""" & vbLf
    s = s & "          onclick=""saveKpiEdit()"">Save Changes</button>" & vbLf
    s = s & "      </div>" & vbLf
    s = s & "    </div>" & vbLf
    s = s & "  </div>" & vbLf
    s = s & "  "
    Sec04_EditKpiModal = s
End Function

'--- Sec05_CreateChartModal : Create Chart modal (type, X/Y column, hierarchy, labels)
Private Function Sec05_CreateChartModal() As String
    Dim s As String
    s = s & "<div id=""createChartModal"" class=""modal-overlay"" style=""display: none;"">" & vbLf
    s = s & "    <div class=""modal-box modal-wide"">" & vbLf
    s = s & "      <div class=""modal-title"">Create New Chart</div>" & vbLf
    s = s & "      <div class=""modal-body"">" & vbLf
    s = s & "        <div class=""input-group"" style=""margin-bottom:15px;"">" & vbLf
    s = s & "          <label class=""input-label"">Chart Title</label>" & vbLf
    s = s & "          <input type=""text"" id=""newChartTitle"" class=""fmt-input"" placeholder=""Enter chart title"">" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "        <div class=""fmt-row"">" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">Chart Type</label>" & vbLf
    s = s & "            <select id=""newChartType"" class=""fmt-input"" onchange=""toggleCreateChartFields()"">" & vbLf
    s = s & "              <option value=""bar"">Bar</option>" & vbLf
    s = s & "              <option value=""column"">Column</option>" & vbLf
    s = s & "              <option value=""line"">Line</option>" & vbLf
    s = s & "              <option value=""area"">Area</option>" & vbLf
    s = s & "              <option value=""pie"">Pie</option>" & vbLf
    s = s & "              <option value=""doughnut"">Doughnut</option>" & vbLf
    s = s & "            </select>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">Aggregation</label>" & vbLf
    s = s & "            <select id=""newChartAgg"" class=""fmt-input"">" & vbLf
    s = s & "              <option value=""sum"">Sum</option>" & vbLf
    s = s & "              <option value=""avg"">Average</option>" & vbLf
    s = s & "              <option value=""count"">Count</option>" & vbLf
    s = s & "            </select>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "        <div class=""fmt-row"">" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">X Column (Category)</label>" & vbLf
    s = s & "            <select id=""newChartXCol"" class=""fmt-input"" onchange=""toggleCreateChartFields()"">" & vbLf
    s = s & "              <option value=""Order_ID"">Order_ID</option>" & vbLf
    s = s & "              <option value=""Order_Date"">Order_Date</option>" & vbLf
    s = s & "              <option value=""Product"">Product</option>" & vbLf
    s = s & "              <option value=""Category"">Category</option>" & vbLf
    s = s & "              <option value=""Region"">Region</option>" & vbLf
    s = s & "              <option value=""City"">City</option>" & vbLf
    s = s & "              <option value=""State"">State</option>" & vbLf
    s = s & "              <option value=""Customer_Name"">Customer_Name</option>" & vbLf
    s = s & "              <option value=""Order_Type"">Order_Type</option>" & vbLf
    s = s & "              <option value=""Customer_Age"">Customer_Age</option>" & vbLf
    s = s & "              <option value=""Quantity"">Quantity</option>" & vbLf
    s = s & "              <option value=""Price_Per_Unit"">Price_Per_Unit</option>" & vbLf
    s = s & "              <option value=""Total_Amount"">Total_Amount</option>" & vbLf
    s = s & "              <option value=""Return"">Return</option>" & vbLf
    s = s & "            </select>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">Y Column (Value)</label>" & vbLf
    s = s & "            <select id=""newChartYCol"" class=""fmt-input"">" & vbLf
    s = s & "              <option value=""Order_ID"">Order_ID</option>" & vbLf
    s = s & "              <option value=""Order_Date"">Order_Date</option>" & vbLf
    s = s & "              <option value=""Product"">Product</option>" & vbLf
    s = s & "              <option value=""Category"">Category</option>" & vbLf
    s = s & "              <option value=""Region"">Region</option>" & vbLf
    s = s & "              <option value=""City"">City</option>" & vbLf
    s = s & "              <option value=""State"">State</option>" & vbLf
    s = s & "              <option value=""Customer_Name"">Customer_Name</option>" & vbLf
    s = s & "              <option value=""Order_Type"">Order_Type</option>" & vbLf
    s = s & "              <option value=""Customer_Age"">Customer_Age</option>" & vbLf
    s = s & "              <option value=""Quantity"">Quantity</option>" & vbLf
    s = s & "              <option value=""Price_Per_Unit"">Price_Per_Unit</option>" & vbLf
    s = s & "              <option value=""Total_Amount"">Total_Amount</option>" & vbLf
    s = s & "              <option value=""Return"">Return</option>" & vbLf
    s = s & "            </select>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "        <div id=""newDateDiv"" style=""display: none; margin-bottom: 15px;"">" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">Date Hierarchy</label>" & vbLf
    s = s & "            <select id=""newChartHierarchy"" class=""fmt-input"">" & vbLf
    s = s & "              <option value=""None"">None</option>" & vbLf
    s = s & "              <option value=""Year"">Year</option>" & vbLf
    s = s & "              <option value=""Quarter"">Quarter</option>" & vbLf
    s = s & "              <option value=""Month"">Month</option>" & vbLf
    s = s & "              <option value=""Day"">Day</option>" & vbLf
    s = s & "            </select>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "        <div id=""newChartNonPieFields"" style=""display: none;"">" & vbLf
    s = s & "          <div class=""fmt-row"">" & vbLf
    s = s & "            <div class=""input-group"">" & vbLf
    s = s & "              <label class=""input-label"">X-Axis Label (Optional)</label>" & vbLf
    s = s & "              <input type=""text"" id=""newChartXLabel"" class=""fmt-input"" placeholder=""Category Name"">" & vbLf
    s = s & "            </div>" & vbLf
    s = s & "            <div class=""input-group"">" & vbLf
    s = s & "              <label class=""input-label"">Y-Axis Label (Optional)</label>" & vbLf
    s = s & "              <input type=""text"" id=""newChartYLabel"" class=""fmt-input"" placeholder=""Value Name"">" & vbLf
    s = s & "            </div>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "          <div class=""input-group"" style=""width: 48%;"">" & vbLf
    s = s & "            <label class=""input-label"">Top N (0 for all)</label>" & vbLf
    s = s & "            <input type=""number"" id=""newChartTop"" class=""fmt-input"" value=""0"" min=""0"">" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "        <div class=""fmt-row size-row"">" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">Width (grid columns)</label>" & vbLf
    s = s & "            <input type=""number"" id=""newChartW"" class=""fmt-input"" value=""6"" min=""1"" max=""12"">" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">Height (rows)</label>" & vbLf
    s = s & "            <input type=""number"" id=""newChartH"" class=""fmt-input"" value=""5"" min=""1"" max=""10"">" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "      </div>" & vbLf
    s = s & "      <div class=""modal-footer"">" & vbLf
    s = s & "        <button class=""btn-secondary"" onclick=""closeModals()"">Cancel</button>" & vbLf
    s = s & "        <button class=""btn-primary"" onclick=""saveNewChart()"">Create Chart</button>" & vbLf
    s = s & "      </div>" & vbLf
    s = s & "    </div>" & vbLf
    s = s & "  </div>" & vbLf
    s = s & "  "
    Sec05_CreateChartModal = s
End Function

'--- Sec06_EditChartModal : Edit Chart modal
Private Function Sec06_EditChartModal() As String
    Dim s As String
    s = s & "<div id=""editChartModal"" class=""modal-overlay"" style=""display: none;"">" & vbLf
    s = s & "    <div class=""modal-box modal-wide"">" & vbLf
    s = s & "      <div class=""modal-title"">Edit Chart Settings</div>" & vbLf
    s = s & "      <div class=""modal-body"">" & vbLf
    s = s & "        <div class=""input-group"" style=""margin-bottom:15px;"">" & vbLf
    s = s & "          <label class=""input-label"">Chart Title</label>" & vbLf
    s = s & "          <input type=""text"" id=""editTitleInput"" class=""fmt-input"" placeholder=""Enter chart title"">" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "        <div class=""fmt-row"">" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">Chart Type</label>" & vbLf
    s = s & "            <select id=""editTypeInput"" class=""fmt-input"" onchange=""toggleEditChartFields()"">" & vbLf
    s = s & "              <option value=""bar"">Bar</option>" & vbLf
    s = s & "              <option value=""column"">Column</option>" & vbLf
    s = s & "              <option value=""line"">Line</option>" & vbLf
    s = s & "              <option value=""area"">Area</option>" & vbLf
    s = s & "              <option value=""pie"">Pie</option>" & vbLf
    s = s & "              <option value=""doughnut"">Doughnut</option>" & vbLf
    s = s & "            </select>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">Aggregation</label>" & vbLf
    s = s & "            <select id=""editAggInput"" class=""fmt-input"">" & vbLf
    s = s & "              <option value=""sum"">Sum</option>" & vbLf
    s = s & "              <option value=""avg"">Average</option>" & vbLf
    s = s & "              <option value=""count"">Count</option>" & vbLf
    s = s & "            </select>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "        <div class=""fmt-row"">" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">X Column (Category)</label>" & vbLf
    s = s & "            <select id=""editXColInput"" class=""fmt-input"" onchange=""toggleEditChartFields()"">" & vbLf
    s = s & "              <option value=""Order_ID"">Order_ID</option>" & vbLf
    s = s & "              <option value=""Order_Date"">Order_Date</option>" & vbLf
    s = s & "              <option value=""Product"">Product</option>" & vbLf
    s = s & "              <option value=""Category"">Category</option>" & vbLf
    s = s & "              <option value=""Region"">Region</option>" & vbLf
    s = s & "              <option value=""City"">City</option>" & vbLf
    s = s & "              <option value=""State"">State</option>" & vbLf
    s = s & "              <option value=""Customer_Name"">Customer_Name</option>" & vbLf
    s = s & "              <option value=""Order_Type"">Order_Type</option>" & vbLf
    s = s & "              <option value=""Customer_Age"">Customer_Age</option>" & vbLf
    s = s & "              <option value=""Quantity"">Quantity</option>" & vbLf
    s = s & "              <option value=""Price_Per_Unit"">Price_Per_Unit</option>" & vbLf
    s = s & "              <option value=""Total_Amount"">Total_Amount</option>" & vbLf
    s = s & "              <option value=""Return"">Return</option>" & vbLf
    s = s & "            </select>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">Y Column (Value)</label>" & vbLf
    s = s & "            <select id=""editYColInput"" class=""fmt-input"">" & vbLf
    s = s & "              <option value=""Order_ID"">Order_ID</option>" & vbLf
    s = s & "              <option value=""Order_Date"">Order_Date</option>" & vbLf
    s = s & "              <option value=""Product"">Product</option>" & vbLf
    s = s & "              <option value=""Category"">Category</option>" & vbLf
    s = s & "              <option value=""Region"">Region</option>" & vbLf
    s = s & "              <option value=""City"">City</option>" & vbLf
    s = s & "              <option value=""State"">State</option>" & vbLf
    s = s & "              <option value=""Customer_Name"">Customer_Name</option>" & vbLf
    s = s & "              <option value=""Order_Type"">Order_Type</option>" & vbLf
    s = s & "              <option value=""Customer_Age"">Customer_Age</option>" & vbLf
    s = s & "              <option value=""Quantity"">Quantity</option>" & vbLf
    s = s & "              <option value=""Price_Per_Unit"">Price_Per_Unit</option>" & vbLf
    s = s & "              <option value=""Total_Amount"">Total_Amount</option>" & vbLf
    s = s & "              <option value=""Return"">Return</option>" & vbLf
    s = s & "            </select>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "        <div id=""editDateDiv"" style=""display: none; margin-bottom: 15px;"">" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">Date Hierarchy</label>" & vbLf
    s = s & "            <select id=""editHierarchyInput"" class=""fmt-input"">" & vbLf
    s = s & "              <option value=""None"">None</option>" & vbLf
    s = s & "              <option value=""Year"">Year</option>" & vbLf
    s = s & "              <option value=""Quarter"">Quarter</option>" & vbLf
    s = s & "              <option value=""Month"">Month</option>" & vbLf
    s = s & "              <option value=""Day"">Day</option>" & vbLf
    s = s & "            </select>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "        <div id=""editNonPieFields"" style=""display: none;"">" & vbLf
    s = s & "          <div class=""fmt-row"">" & vbLf
    s = s & "            <div class=""input-group"">" & vbLf
    s = s & "              <label class=""input-label"">X-Axis Label</label>" & vbLf
    s = s & "              <input type=""text"" id=""editXLabelInput"" class=""fmt-input"" placeholder=""Category Name"">" & vbLf
    s = s & "            </div>" & vbLf
    s = s & "            <div class=""input-group"">" & vbLf
    s = s & "              <label class=""input-label"">Y-Axis Label</label>" & vbLf
    s = s & "              <input type=""text"" id=""editYLabelInput"" class=""fmt-input"" placeholder=""Value Name"">" & vbLf
    s = s & "            </div>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "          <div class=""input-group"" style=""width: 48%;"">" & vbLf
    s = s & "            <label class=""input-label"">Top N (0 for all)</label>" & vbLf
    s = s & "            <input type=""number"" id=""editTopInput"" class=""fmt-input"" value=""0"" min=""0"">" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "        <div class=""fmt-row size-row"">" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">Width (grid columns)</label>" & vbLf
    s = s & "            <input type=""number"" id=""editChartW"" class=""fmt-input"" value=""6"" min=""1"" max=""12"">" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "          <div class=""input-group"">" & vbLf
    s = s & "            <label class=""input-label"">Height (rows)</label>" & vbLf
    s = s & "            <input type=""number"" id=""editChartH"" class=""fmt-input"" value=""5"" min=""1"" max=""10"">" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "      </div>" & vbLf
    s = s & "      <div class=""modal-footer"">" & vbLf
    s = s & "        <button class=""btn-secondary"" onclick=""closeModals()"">Cancel</button>" & vbLf
    s = s & "        <button class=""btn-primary"" onclick=""saveChartEdit()"">Save Changes</button>" & vbLf
    s = s & "      </div>" & vbLf
    s = s & "    </div>" & vbLf
    s = s & "  </div>" & vbLf
    s = s & "  "
    Sec06_EditChartModal = s
End Function

'--- Sec07_LayoutModal : Layout picker modal - 6 preset layouts + custom KPI/chart/slicer layout
Private Function Sec07_LayoutModal() As String
    Dim s As String
    s = s & "<div id=""layoutModal"" class=""modal-overlay"" style=""display: none;"">" & vbLf
    s = s & "    <div class=""modal-box modal-wide"" style=""width: 750px;"">" & vbLf
    s = s & "      <div class=""modal-title"">Professional Layout Builder</div>" & vbLf
    s = s & "      <div class=""modal-body"">" & vbLf
    s = s & "        <div class=""layout-sections"">" & vbLf
    s = s & "          <div>" & vbLf
    s = s & "            <div class=""input-label"" style=""margin-bottom:8px;"">1. Create Custom Layout</div>" & vbLf
    s = s & "            <div class=""custom-builder-card"">" & vbLf
    s = s & "              <div class=""input-group""><label class=""input-label"">KPI Cards</label><input type=""number"" id=""layKpi""" & vbLf
    s = s & "                  class=""fmt-input"" value=""5"" min=""0"" max=""12""></div>" & vbLf
    s = s & "              <div class=""input-group""><label class=""input-label"">Charts</label><input type=""number"" id=""layCharts""" & vbLf
    s = s & "                  class=""fmt-input"" value=""2"" min=""0"" max=""10""></div>" & vbLf
    s = s & "              <div class=""input-group""><label class=""input-label"">Tables</label><input type=""number"" id=""layTables""" & vbLf
    s = s & "                  class=""fmt-input"" value=""1"" min=""0"" max=""6""></div>" & vbLf
    s = s & "              <div class=""input-group""><label class=""input-label"">Slicers</label><input type=""number"" id=""laySlicers""" & vbLf
    s = s & "                  class=""fmt-input"" value=""5"" min=""0"" max=""5""></div>" & vbLf
    s = s & "              <button class=""btn-generate"" onclick=""applyCustomLayout()"">Build Layout</button>" & vbLf
    s = s & "            </div>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "          <div>" & vbLf
    s = s & "            <div class=""input-label"" style=""margin-bottom:8px;"">2. Or Choose a Pro Template</div>" & vbLf
    s = s & "            <div class=""layout-grid"" style=""grid-template-columns: repeat(3, 1fr);"">" & vbLf
    s = s & "              <div class=""layout-option"" onclick=""applyLayout(1)"">" & vbLf
    s = s & "                <div" & vbLf
    s = s & "                  style=""height:40px;background:linear-gradient(to bottom, #cbd5e1 40%, #e2e8f0 40%); border-radius:4px; margin-bottom:5px;"">" & vbLf
    s = s & "                </div><strong>Balanced Classic</strong><br><span style=""font-size:10px;"">4 KPIs + 2 Charts</span>" & vbLf
    s = s & "              </div>" & vbLf
    s = s & "              <div class=""layout-option"" onclick=""applyLayout(2)"">" & vbLf
    s = s & "                <div" & vbLf
    s = s & "                  style=""height:40px;background:linear-gradient(to right, #e2e8f0 30%, #cbd5e1 30%); border-radius:4px; margin-bottom:5px;"">" & vbLf
    s = s & "                </div><strong>Hero Chart</strong><br><span style=""font-size:10px;"">3 KPIs + 1 Main + 2 Sub</span>" & vbLf
    s = s & "              </div>" & vbLf
    s = s & "              <div class=""layout-option"" onclick=""applyLayout(3)"">" & vbLf
    s = s & "                <div" & vbLf
    s = s & "                  style=""height:40px;background:linear-gradient(to right, #cbd5e1 70%, #e2e8f0 70%); border-radius:4px; margin-bottom:5px;"">" & vbLf
    s = s & "                </div><strong>Sidebar Focus</strong><br><span style=""font-size:10px;"">1 Large Chart + Side KPIs</span>" & vbLf
    s = s & "              </div>" & vbLf
    s = s & "              <div class=""layout-option"" onclick=""applyLayout(4)"">" & vbLf
    s = s & "                <div" & vbLf
    s = s & "                  style=""height:40px;background:linear-gradient(to bottom, #e2e8f0 30%, #cbd5e1 30%); border-radius:4px; margin-bottom:5px;"">" & vbLf
    s = s & "                </div><strong>Split Panels</strong><br><span style=""font-size:10px;"">2 Wide KPIs + 2 Wide Charts</span>" & vbLf
    s = s & "              </div>" & vbLf
    s = s & "              <div class=""layout-option"" onclick=""applyLayout(5)""" & vbLf
    s = s & "                style=""border-color:#10b981; box-shadow:0 0 10px rgba(16,185,129,0.1);"">" & vbLf
    s = s & "                <div" & vbLf
    s = s & "                  style=""height:40px;background:repeating-linear-gradient(45deg, #10b981 0px, #10b981 5px, #a7f3d0 5px, #a7f3d0 10px); border-radius:4px; margin-bottom:5px;"">" & vbLf
    s = s & "                </div><strong>Executive Grid</strong><br><span style=""font-size:10px;color:#10b981;font-weight:700;"">6" & vbLf
    s = s & "                  KPIs + 4 Charts</span>" & vbLf
    s = s & "              </div>" & vbLf
    s = s & "              <div class=""layout-option"" onclick=""applyLayout(6)"">" & vbLf
    s = s & "                <div" & vbLf
    s = s & "                  style=""height:40px;background:linear-gradient(135deg, #e2e8f0 50%, #cbd5e1 50%); border-radius:4px; margin-bottom:5px;"">" & vbLf
    s = s & "                </div><strong>Analytics Pro</strong><br><span style=""font-size:10px;"">6 KPIs + 3 Charts</span>" & vbLf
    s = s & "              </div>" & vbLf
    s = s & "            </div>" & vbLf
    s = s & "          </div>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "      </div>" & vbLf
    s = s & "      <div class=""modal-footer""><button class=""btn-secondary"" onclick=""closeModals()"">Cancel</button></div>" & vbLf
    s = s & "    </div>" & vbLf
    s = s & "  </div>" & vbLf
    s = s & "" & vbLf
    s = s & "  <div id=""dataViewModal"" class=""modal-overlay"" style=""display: none;"">" & vbLf
    s = s & "    <div class=""modal-box modal-wide"" style=""width: 900px;"">" & vbLf
    s = s & "      <div class=""modal-title""><i class=""fa-solid fa-table-list""" & vbLf
    s = s & "          style=""color:var(--kpi-val-col); margin-right:8px;""></i>Raw Data View (Top 50 Rows)</div>" & vbLf
    s = s & "      <div class=""modal-body"">" & vbLf
    s = s & "        <div class=""dv-tabs"" id=""dvTabs""></div>" & vbLf
    s = s & "        <div class=""dv-wrap"" id=""dvWrap""></div>" & vbLf
    s = s & "        <div style=""font-size:11px; color:var(--axis-col); margin-top:8px;"" id=""dvMeta""></div>" & vbLf
    s = s & "      </div>" & vbLf
    s = s & "      <div class=""modal-footer"">" & vbLf
    s = s & "        <button class=""btn-secondary"" onclick=""closeModals()"">Close</button>" & vbLf
    s = s & "      </div>" & vbLf
    s = s & "    </div>" & vbLf
    s = s & "  </div>" & vbLf
    s = s & "  <div id=""dataModelerModal"" class=""modal-overlay"" style=""display: none;"">" & vbLf
    s = s & "    <div class=""modal-box modal-wide"" style=""width: 940px;"">" & vbLf
    s = s & "      <div class=""modal-title""><i class=""fa-solid fa-diagram-project""" & vbLf
    s = s & "          style=""color:var(--kpi-val-col); margin-right:8px;""></i>Data Modeler - Table Relationships</div>" & vbLf
    s = s & "      <div class=""modal-body"">" & vbLf
    s = s & "        <div class=""dm-hint"">Click a column, then click a column in another table to draw a link. Drag a table" & vbLf
    s = s & "          header to move it around. Relations are saved with the dashboard.</div>" & vbLf
    s = s & "        <div class=""dm-canvas"" id=""dmCanvas""><svg class=""dm-svg"" id=""dmSvg"" xmlns=""http://www.w3.org/2000/svg""></svg>" & vbLf
    s = s & "        </div>" & vbLf
    s = s & "        <div class=""dm-rels"" id=""dmRels""></div>" & vbLf
    s = s & "      </div>" & vbLf
    s = s & "      <div class=""modal-footer"">" & vbLf
    s = s & "        <button class=""btn-secondary"" onclick=""closeModals()"">Cancel</button>" & vbLf
    s = s & "        <button class=""btn-primary"" onclick=""saveDataModel()"">Save Relationships</button>" & vbLf
    s = s & "      </div>" & vbLf
    s = s & "    </div>" & vbLf
    s = s & "  </div>" & vbLf
    s = s & "" & vbLf
    s = s & "  "
    Sec07_LayoutModal = s
End Function

'--- Sec08_HeaderLogo_1 : Dashboard header, toolbar buttons (Layout/Update/KPI/Chart/Slicer/Save/Reset/Deploy) and the embedded company logo bitmap (chunk 1 of 18)
Private Function Sec08_HeaderLogo_1() As String
    Dim s As String
    s = s & "<div class=""dash-header"">" & vbLf
    s = s & "    <div class=""header-top"">" & vbLf
    s = s & "      <div style=""display: flex; align-items: center; gap: 15px;"">" & vbLf
    s = s & "        <div id=""dashLogo""" & vbLf
    s = s & "          style=""width: 45px; height: 45px; display: flex; justify-content: center; align-items: center; cursor: pointer; border-radius: 8px; transition: 0.3s; border: none; background: transparent;""" & vbLf
    s = s & "          onclick=""openLogoModal()"" title=""Click to Add Company Logo""><img" & vbLf
    s = s & ""
    Sec08_HeaderLogo_1 = s
End Function

'--- Sec08_HeaderLogo_2 : Dashboard header, toolbar buttons (Layout/Update/KPI/Chart/Slicer/Save/Reset/Deploy) and the embedded company logo bitmap (chunk 2 of 18)
Private Function Sec08_HeaderLogo_2() As String
    Dim s As String
    s = s & "            src=""data:image/x-icon;base64,AAABAAkAEBAAAAEAIABoBAAAlgAAABgYAAABACAAiAkAAP4EAAAgIAAAAQAgAKgQAACGDgAAMDAAAAEAIACoJQAALh8AAEBAAAABACAAKEIAANZEAABISAAAAQAgAIhUAAD+hgAAYGAAAAEAIAColAAAhtsAAICAAAABACAAKAgBAC5wAQAAAAAAAQAgALa7AABWeAIAKAAAABAAAAAgAAAAAQAgAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAammxAGRhuQA9Q1oAQIzeAGH//wI2e9KUJVWdqRcsXAgjO2sALDFEAGFduwBgXb0AX1u+AAAAAAAAAAAAeGj0AHtq+gBhWeUAWFTJANPs/wBBbp8iPnrI2CNJhuwXJkM6a3mUAEVEkAA+PbMAPz6yAD8+sQAAAAAAAAAAAHpp8wB7a/kCYFniJU5MzwjRx8oplJGWunV3n/xPUHn/TUxQz2ViaUJJSKcMPj2xOj8+sQg/PrAARESUAAAAAAB3aPUAeGn2FGhb7cFvZOWj6+Duye3g5v++suD/e3ap/4iGiP+BfofhRkWesDs7rt09PLAtPTuwAExLnAAAAAAAeWn3AHhp9hprXfHbe2/u/+nd8//16O3/x7nq/355s/+OjI//iYeP/01Mof85OKzxOzqtNTs6rgBJSJkAAAAAAIFv+QCSfP8DbV70kIl88f/05vX/9unu/8/A7f+Jgrv/lJKV/5GQk/9cW6D/PDuxrEE+sg0/PbEA////AAAAAAB/bfoAc2T2AFZG9hGzpfWR+ez2/vXn7f/Yy93/rKW7/6ajpf+ZmJj/e3qauD48uxs/PbMAODauAGdktgAAAAAAgXvdAG1e9gDg0vYA//b2Uf/x9fze0dn/loyt/4F5o/+Cf4//nJqb"
    s = s & "/5eWl6EAAP8AurmXAFRStwAAAAAAAAAAAAAAAADy7esA/vH1AP7y9l397/L+pJm9/6Sa3//Iv/T/c2qz/5OQlf+cm5y2jI2VB5SUmwAAAAAAAAAAAAAAAAAAAAAA7ujrAP/x9gD/8vdh/O/z/5+Uw/+zqO7/3db8/31yxv+Wkpn/o6GhwpOUlguXl5oAAAAAAAAAAAAAAAAAAAAAAOni7AD/8fcA/vH3Xf/y+P7Xytr/mY7F/5aLy/+Mhab/r6ut/6yoqcaenZ8NoZ+hAAAAAAAAAAAAAAAAAAAAAADHwesA9uj2APXo9UL56/n2/O72/+TW3//MwMv/y8LI/764wP+sqLOzkpOcBqCfpQAAAAAAAAAAAAAAAAAAAAAA0crwALCe8gCunPEYr5r4zrWi9/+yofT/ppfv/5GF4/91btH/ZmLEdm5n5QBeXa4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8ArpT8AI1x/V6Sd/73jXb+/4Fw//9uYvX/XFTmzk9M3h1VUdwAZGObAAAAAAAAAAAAAAAAAAAAAAAAAAAAiXX6AI10/ACKc/oGk3f8j5F3/f6Bcf7/bWH07GBX501jVugAAAByAHt5xwAAAAAAAAAAAAAAAAAAAAAAAAAAAGpotgCPdP4AkHX8AJB1+hOQdf2jf2396Gxg83JESMQBX1rmAFpY4gAAAAAAAAAAAAAAAACAAQAAgAEAAIAAAACAAAAAgAAAAIAAAACAAAAAgAEAAMADAADAAwAAwAMAAMADAADAAwAA4AMAAOADAADgBwAAKAAAABgAAAAwAAAAAQAgAAAAAAAACQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADdAD6P5gBCmfANNnjQqCVUnMAdOmwcHD5xAK2H/wAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABDSIcAUGQAAEKK1wBDmPFEOYHZ9CRXn/8cPHNtFjhoADdGgAAzP2UAZGG5AGFdvABhXbwAYV68AAAAAAAAAAAAAAAAAAAAAAAAAAAAZ2W1AGlg6ABiXuAATFPMACw5jQAAAAAAxbmzAC8DAAVMiM18P3vM/yRMjf8aMFeuIyg4FP/SaAA+SXoAIyegADo4rAA6OK0AOTetAF9cvQAAAAAAAAAAAAAAAAAAAAAAlZTfAHBm5gBgXtgBQ0+/Bl9U4wC7tMgAbm5/CmZma5VrbnXqWl2O/z9Bd/80NDr1MjA3uEFBTiFWU0YANjWfADs5rRM7OK0KOziuAAgMeQAAAAAAAAAAAAAAAAAAAAAAionSAG9h8wByZPMxYVjmoVFL0DXz6e4c7OHmndLIzvTNw8r/oZjO/21pqv+DgIb/dnN2/Hh1e8d1c39ENzajOT49sr4/PrJmPT24AExMrAAAAAAAAAAAAAAAAAAAAAAAgnvVAGhc8gByZPNraFvu/2BV5d+uo+m3+e7y/vXo7v/r3ev/tqrn/21os/+IhY7/i4mM/4mHiv9saY7YOTik4zo6rf88O7CqPTmpBD06rAAAAAAAAAAAAAAAAAAAAAAAdGraAGRW8gBzZPR7alzw/2Za7f+vo+//++/z//Xo7f/r3er/u67r/3JtvP+IhpH/j4yP/46Mjf9xb5L/Ozum/zk5q/87Oq23PDiqCDs4qgAAAAAAAAAAAAAAAAAAAAAAfXrSAHJi9gB4Z/dQbV7z+Wpc7v/CtfP//vH1//bp7v/t3+r/w7Lv/3t0xv+LiZb/ko+S/5GPkf99e5X/Q0Gm/zo6rf89O6+GRUTjADo6nAAAAAAAAAAAAAAAAAAAAAAAiorWAH9t+QCDcPoP"
    s = s & "cWP1rG5h8P/PxfX///L1//jq8P/u4Ov/yLnw/4J5yf+QjZj/l5WX/5WUlf+IhpP/TUuk/z49ss1EQrYmQ0G1AHFrjAAAAAAAAAAAAAAAAAAAAAAAdHHLAHlk/AB0Y/gAdGX5HHts85Lg0/f0//L2//fp7//t3+j/1cjl/6mfzf+ppKv/o6Cj/5qYmv+Qj5P/Y2GjskA+viowK64AscLeAEZGkwAAAAAAAAAAAAAAAAAAAAAAAAAAAIiB6wBmVvgAr6H3APLm9hb87/bT/fD1//bo7f/m2d3/wbi8/6iip/+loaP/p6Sl/52cnP+VlJb/kZCUanZ0nACireQAT1O9AF9bvgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB6eMkA//P3AP7z9xv/8fXZ/u/0/+7g5P+dlKn/cGef/25lpv9hWov/dnV//6Cenv+bmpr/lJOWg5ubigCFgr0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA/fL3AP3y9yH+8fXh//H1/8K3wv9yZrL/ubDv/9TM+f+bkOj/XFWO/5SQk/+fnp7/lpWXmUdG1gGRkKUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA/vP4AP7z+CX/8fbl//L1/6WcsP92adr/z8b4//Hr+v+tofj/a166/4uIjP+lo6P/m5qbpoeNkgOVlpsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA//P4AP7z+CX+8fbl//L4/7Wrvv54a9H/0sn4//Ls+/+uo/b/al6v/5OPlP+ppqf/nZydtIyNlgaUlJ0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA/vP5AP7y+CT+8Pfk//H5/+fa4v+Jf6//lInW/6KX5v+CeMb/eHOP/6"
    s = s & "2prP+uqqv/pqOkuZeXnQednaIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA/fL3APzx9hr+8ffX/vD4///x9//g09r/ppq1/4+Eqv+clKT/urS3/765vP+2sbT/raqrrZSXmQShoKUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA8ebwAOvh7Qj26fe2/e/6//7x+f/97vT/9ufr/+rc3v/h1dn/1MrP/8jBxv+5tLv/pqOwjd/P1wCZmqQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAj4XJANuz/wCtmvN+t6T4/8Kx9v/Ds/L/v6/w/7Sn6f+onOT/mIzY/4F6yv9rZ8T6XVu3WGJfvgBra9EAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//+dAIxz+wCFbPsjkHX81ZR5/f+PePz/iHX8/4Nx/P97bPr/bGDw/11U5P9UUNy/TEnTEk1K0QCHhNgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAdHLLAIBy8wCYd/8AjnT9YJR5/vaVev7/j3f+/4V0//97bf3/bGDy/15W5vFXVOBMVlHjABkW6QB+fcsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIhv/wCLcvwAh3L4B5J3/JCUeP3/kHf9/4Z0//94a/z/al7x/2BX54hPVtkEWFbhAFZTygAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wCNcP8AkHT9AI50+hSSdvyokXf9/4Ry/v9zZvn/Z1zvrF5Y5BNeV+gAWFPsAGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAACOjNsAlnb/AJJ2/QCPdPsbj3X9ooNv/uNyZfisZlvsIWJY7ABWTu4AY2KwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8A/wD+AAcAwAADAMAAAwDAAAMAwAADAMAAAwDAAAMAwAADAMAAAwDgAAMA8AAPAPAADwDwAA8A8AAPAPAADwDwAA8A+AAPAPgADwD4AA8A+AAPAPwAHwD8AB8A/gA/ACgAAAAgAAAAQAAAAAEAIAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD/aWMANY7nAD2R6Bg2dc+zJVOcyhs9cTEPPnEAAAAAAF1cqgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAdXPBADyP7AA7h9sAP4/nbzZ81P0mWqX/HUB4ojE2ZwcnO24AU1KgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFFTmgBKSwoARYXIAEul/wtFl++8OYDZ/yddqv8aPnfmIDprKyE4XwBWXsUAMD1kAGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AGhmtABnZLQAa2jCAIOB1QC9ueMAEAkAAF9dUQAmAwACTG+kJ02M1tc8dcf/JlCX/xoyXPgeL01eMzdMCwAAAABug/8Avb/jACgpnAA2Np0AOj"
    s = s & "emAAMKcABiXrsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABMQvkAXVbtAHZk+QBeVOYAQkC6ALux4QDFvsgAAAAAAE9RWXdeYGXTWl5r905RiP87PX3/KSk0/SQkKecmJi6oISU0EVxaWgCEhogAOjqiAAAAAAA3MpoDNzenADk3sABfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEqZABdWeoAamDrNVRS11A0NpcFfnfWAO3j5xDc0dhVpJ+j2aagpP+ooqf/hH27/1xYo/91cnz/ZWNm/1tXW/NoZm2MbWx6MxgbkABAQZ4HPj2yaj08sntAP6sGPTysAFhVtgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABcV8IAhXH6AJeC/AVvYfKlY1jp9VpT2IxVUcYe+O3ygfTo7Prx5Ov/7d/q/+DT5P+nnOH/ZGC1/4yJmf+Mio3/h4SH/4J/hP98eoLMT06MOzk3opQ9PbD3Pz+z5T8+rSo9PK4AUlCkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGVitwB/bfkAgXD2Fm5h8dJnWu7/ZFjq/2ti4NHj1/Ll+u7x//Xo7f/v4ev/49Tp/66i6f9kX7v/gX6S/4qIi/+LiYz/i4mL/4J/jPpNSZbkODeo/zk5q/87Oq/6PDqvVDk4sgBNTJ8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYF29AH9v/AB+b/ggbmHy3mhb7v9nWu//eG3p/+PX9f/77/L/9unu/+/h6//j1en/s6bs/2pkw/+Bf5b/jYqO/4+Lj/+NjI3/h4WK/1VSmv83N6n/ODiq/zo6rf07Oa5iOTm0AEREkwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABjYLkAhXT9AIh3/BFyYvXIa1vy/2ha8P+FeOz/7+H3//zv"
    s = s & "8//26e7/8OLr/+bY6v+5qe//c2vM/4OBnP+PjJD/kI6R/5CNkf+KiI3/YF6c/zk4qP85Oaz/Ozqt9Ds5qkQ5N60AVVSoAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGNguQCNefkAAADMAHhn94BtXvP/aFrw/5KG7f/36vj//u/0//jr7//x5Oz/6Nnq/76t8/97cdP/hoSf/5KPk/+TkZT/k5GU/4+NkP9vbJ3/PT2o/zo6rv8+PbLAQj6yE0A8tABOTpoAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALKM/wB/bvwAgXD6HnFk9cFpWvH/npTw//vv+P/+7/X/+evx//Hk7P/p2ur/w7Ty/39z1f+Khp3/lpOW/5eVmP+VlJb/kpCS/3t5lP9DQaj/Pz6z4kdFt0ZDQ7YAdmvCAENDkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX1u+AHJZ/wBxYPoAeWr5KG1d9KStoPTy/fD4//7w9//56vH/8OLr/+jb5//Owen/m5DV/6Kbrf+loaX/oZ6h/5qZmv+VlJX/hoWT/VBOrLxFQr0/////AFJUwQAuKKoAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhXn2AP//5AB0Y/kAXUrwCOba9pb+8fb//fH1//fq7//w4uf/5tnf/9nO1f/Jwcn/vbm8/7KvsP+npab/nZyc/5aWl/+PjpPygoCeQGppoQBXWcQASk/DAF9bvgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAi4rqAPv18QC+jvUA//P3jf/x9v/+8PT/9ujt/+ve4v/EusL/k4yV/394g/+Ae4T/k5CS/6Shov+gnp7/mpmZ/5WUlfiUlJRMk5OQAI+O1gCOjdsAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA9+31AOP1/wH+8faa//D1//7u9P/15+z/urC3/2pgjP9pX6v/b2S3/2Zbqf9LSW//dHR7/6Gfn/+fnp7/mJeY/5KRlmOXlowAZGG7AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wD06/YA8fD5Av3x9ab/8fX//e/0/+HU2f9xaI7/gXTQ/8K58//Kwff/vrP0/3No0f9STnH/lpKV/6KgoP+ZmZn/kpKWe5ycegCEgs0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7ujrAPrw9wD79PwE/vL3rv/x9f//8vb/vrS5/2pfrv+Lfe7/5uH6//Lt/P/d1/r/g3Xx/2FWnP+GgoX/pqSl/5+dnv+YmJiHoZx9AIeGyQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA/vP3AP31+gX+8vew//H1///z+P+1rLT/b2O6/4l79P/f1/j/7OX4/9bO+P+DdPf/alyw/4SAhf+opqf/oZ+g/5qZmZf//70AhIO/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wD+8/gA/fb6Bf7x97D+8Pf///P5/87Fzf5yZq7+jH3r/+nj+v/z7fr/3dj6/4Z37f9jWZb/kIyR/62pqv+koqP/m5qbpXJ0owKMjrkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7ujrAP3y9wD89foF/vH4r/7w9//+8fj/9eju/5SJpf96b8H/q6Dt"
    s = s & "/7Cm8/+mm+3/cGS7/25qf/+tqKz/r6us/6qnqP+joKGpiY6fA5mhpQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA+e73APf2/wL+8fej/vD3//7w+P//8fj/5trf/5eOpP9zaKb/bmKw/21jnP98d4r/rams/7q1uf+1sLP/r6ut/6mmp6N8hpsClZqvAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wDj2/AA////AP3x9YT+8Pf///H5///x+P/87fP/8uPo/9XIzf+/s7r/w7i9/87Fx//Jw8X/w77A/725u/+2srT/rKqqi/ro1gCUktAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJqY3QD77vkA8ub1VPrt+fr/8vv///L6//7w9v/77PL/9+js//Hi5f/n2t7/3tPX/9TLz//Nxcr/w73D/7OvuP+fna5mpqKyAKqo1gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAfHrRAKuZ7gCjk+YutaH05cOy9//Nv/b/0cTz/9LD8f/Nv+//xLfp/7uv4v+wpOD/oZbS/4yFxP94c8H/ZGC57lVTqz1WU6sAg4LSAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAfWX3AGxa7ASMcvqRknf7/5V7+f+UfPj/jXr2/4l39v+Ec/X/f2/0/3dp8P9rYOf/Xlbf/1dR3P9RTdaoTUzPCEpH0AB0c8EAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAHRyywCMUP8AjXL+AItz/SeRdf3YlXr//5R6//+PeP7/inb//4Vz//+CcP//eGr7/2tf8f9gVuj/WFPg6lVS2D5WUeIALyz/AGlntQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJiZ7ACIdvYAnHb/AI10/WGTeP71lnv//5R4/f+Od/3/hnT//39x//91aPr/al7w/2BW5/9YVOKHXmnbAlhY3ABSTtIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX1u+AIti/wCKcf0AhnH4B5F2/I+UeP3/k3f8/493/f+Hdf//fnD+/3Rn+v9pXe//YFfov1hW4RlVVOIANzfjAI6N2wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjo3bAItl/wCPcv8AjXT5FpJ2/LGTd/z/kHf9/4d0//97bf3/cGT3/2Zb7d5eV+U5XFXwAC8o/wCIh90AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjIvZAKN4/wCSdv0AjnT7IZJ2/LaQdv3/h3L//3dp/P9rX/PnZFrrUWJG/wBhYdwAU03LAGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAfGjxAKR//wCTd/8AjnT7IY90/pyGcP7YeGn8w21g8lNqccYB"
    s = s & "ZlzpAFJH8wBxb70AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//AH///gB///wAH/8AAAD/AAAAfwAAAH4AAAB+AAAAfgAAAH4AAAB+AAAAfwAAAH8AAAD/gAAA/8AAAf/gAAP/wAAD/8AAA//AAAP/wAAD/8AAA//AAAP/wAAD/+AAA//gAAP/4AAD/+AAA//wAAf/8AAH//gAD//8AA///gAf8oAAAAMAAAAGAAAAABACAAAAAAAAAkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wD//74AKIPZADeF3y84b864J1KezBg/dVVONHEBABxOAH12zQBfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX1u+AJuF0wAnkOgARZTnDzyL4600d87/KF2r/xxFgdwlO2w3BDg+AEI5mABXUccAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYV2/AEe5/wA5me4AQJPrTz2J4fQ0d87/K2K0/xxJiP8eO3CfNzxxBw"
    s = s & "cmRQBiX7UAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAdXPKACGm/gA+mPEFQ5btmkGN5f80d8//LGO0/x1LjP8aO3DfKT9oKh07WABLS5gAZGKwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGhmtABoZroAAAAAADyI0QBfn+8QS571ykeY8P83ftf/K2S1/xxIh/8XNWj4HTpsVRk5WgAAAAAAAAAAAGlntQBoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAaGa0AGVjtwBgXbwAYl+7AAAAAAAAAAAA7ujrAGVhrwAAAAAAV29oACwqVAVCX5wkT4rSz0iCzv81abv/KVOf/hw6a/4aLE/+HC5KdjM7TxoySocECgwAACwsAAAOEV0AaGa0AAAAAABdWrsAWle3AGFeuwBgXL0AYFy9AF9bvgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGhmtABgaHsAgGn/AExT2AAwS64AsKLjAKKi8ADq"
    s = s & "5OkA7efqAKOh1QD/KAAAQkdXQldYXKlaXGTOXF9r9ExNZP9MTo7/QEGS/ygqQf8mJi3+IiIp5iQkK9MoKziGMTZJDRUdAABKSKgAfoGpAA0PWwD///MAEBWYACwsngAtKqUAKyarACYyYwBfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AG5tqgCMdugAsX7/ADNOsAQAMwAAHiShAE9bSQD///8A//f9ANnP1QBFSZkDWFlgpWNjZv9wcHH/cXFy/2Fhbf9hXKL/SEWf/z48U/86OT3+MjE0/y4rL/8uLDLfOzpCLADl/wCEfloAU2r/AFBC2wAzM6UAMzGWAzg2pSs7N6sUNjCyADod/wBfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF5etgCGcP8Afmv2Il9a4pdJSspWLSyKBExH2gD68fIA+fH2BOHX3Dvi1t5+vba517y0uf+/uLz/xLu+/7ixxP+KgdL/U1Cv/4iFpP+Oi5D/f36A/3Nwcv9qZmr4b2txxGdndIZybYErX1RdAD41tABDRp8HQD6yaD07s98+PLKHPDygBCwsnABrZtIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhYTTAGIk/wAkDf8AdGb1cmpd8P9dVePtVlDOgUFBrhiaj+kA/PL2RvPo7eLy5ur/8eTr//Dj6//s3+n/4dTh/9HG4f+XjOD/VFG4/4SAo/+RjZj/jouO/4eFh/+EgoX/gX6D/316fv95dn22cXR7EDczmx45N56OPj+w8j9Atv9AQLPdQECsJzo6sgBOTKEAZGG4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjYvPAIBk/wCbg/4JdWb0rmhb8P9mWe3/Ylfm/VtU2LxWUclM8ub2mfnu8v/47fD/9ejt//Hk6//t3+v/5dbo/9fJ6v+ekuX/V1K+/3Jwnf+Gg4//ioiL/4qIi/+KiIv/ioiL/4eEif9+fIzuTU2IgDg2ncU5OKf/Ojqs/zw8sP89PbH3PDusUzQzpwAAAGoAame+AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAZ2W3AIRw/wCBcfMVcWTz0Ghb7/9nWu7/Zlnt/2VZ6v9rYt7y1sry9Pvv8//47fH/9unu//Ll6//t3+v/5dbo/9fK6v+iluv/XFbG/3Nwn/+Gg4//ioiL/4uIi/+LiIz/i4mM/4iHiP+CfpD+VVCT/Dg3pP84OKr/OTmq/zk5q/87OrD/OzmvfldOSwGyuVkATEueAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX1y9AJB5/wCCcfkjcWTz3Whb7v9oW+7/aFvv/2db7f92bOb/3tL3//vv8//57vL/9+vv//Pm7P/u4Ov/5dbo/9fK6v+nmu3/Y1zP/3Jwof+Gg5L/jImN/46Kjv+Pi4//joyO/4qLif+GhIr/X1uU/zg3pv84OKn/ODip/zk5q/86Oq7/OzmvikM4nANddQAAQkKPAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYl67AJB+/wCFd/sZc2X21mlc8P9pW+//aVvw/2Za7f+Bduf/6Nz6//vv8//67vP/9+vv//Pm6//u4ev/5tjo/9rL6/+pm+7/aF/U/3Nxp/+Gg5X/jouP/4+Nj/+PjY//j42Q/42MjP+Ihov/ZmOW/zo6"
    s = s & "pv84OKn/OTmq/zo6rP87O6z/OjmrfF9KAAC5u34ATUufAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAcW/BAJR9/wCbhv0JeGf3rmxd8/9sXPP/a1zx/2ZZ7f+Rg+v/8uP7//3v9P/87/P/9+vu//Pm6//v4ev/6dvp/97O7f+unfD/cmfa/3d1rv+Kh5n/j4yQ/5CNkP+RjpL/kI2S/46Ljv+Jh43/cW6c/0E/ov84OKr/OTmr/zo6rf87Oq/1OTioTjUysQAJFE0Aa2i/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAiYjZAIdb/wBbQv8Afm34Z29g9PtrW/L/a1zy/2ZZ7P+ekuz/+Ov8//7v9P/97/T/+Ozw//Tn7P/w4uz/6dvp/97O7f+zofX/eWzh/3l2r/+LiZr/kY6S/5KPk/+TkJT/k5CU/5COkP+NjI3/e3mf/0hGoP86Oqv/Ojqt/zo6rv9APrTIQT2vGjUwuwBOTZoAZWK3AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAiojYAGZosACJcv8AiHL7GXdo+MRrX/H/a13y/2hb7P+qoe//+vD6//7v9P/+7/T/+e3x//Xo7f/x4+v/6dvp/97P7P+1pPb/fG7j/3l3r/+MiZr/lJGU/5SSlP+VlJb/lZSW/5OSlP+QjpH/g4Ca/1NPnv86Oqz/Ojqt/z8+sfZDQLVdNTrQADst+ABbWqcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFxYvQC9qf8AZFP/AH5v+kJyZPbealvy/2pd7v+1rfP//PL6///w9f/+7/X/++zy//Xn7f/v4u"
    s = s & "z/69zq/+PU6v+9rvT/gHHl/396rf+RjJj/l5OX/5iVmP+Xlpj/lpWX/5OTlP+QjpH/iYiO/1xZmf88Oq//QD+090pJuYhXVb8KRUO2AGdktABKSZYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF9bvgBvg5oAgmT/APLG8QF5avlFcWH1xW1e8P2/s/b//vL6//7w9//+7/f/++v0//Tm7f/u4ez/6tzp/+PV5//EtvD/iXrj/42Gtf+elqT/op2j/6Kdov+fnJ//m5ib/5aVlv+UkpT/j4+P/2dlm/5BQLXeRkO7cFtdyAxgWeEALUVFAF9bvgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACOjdsApVj/AJqF/wBhO/8AgW77G3lp8HnXyPbl//L3//7y9//98Pb/+ezw//Tm7P/v4ej/6dvk/+DU3v/Sxtz/v7TY/7mzw/+3srf/sK2v/6unqf+loqT/npyd/5mZmf+VlJb/kI+R/3l3n8ZOS8EyeXXzATU3xAAAAP8AQkKQAF9bvgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX1u+ADUtZQD///8Aqpf/APPn7Q347vi4/vL2//7y9v/88PX/9+rv//Tm6//v4eT/59re/97S1//YzdD/zsfJ/8W/wv+8uLr/tLCy/62pq/+lo6T/nZyc/5mZmf+VlZb/kpGS/42Lk7GNjaIKg4KTAJyd0wCcnOQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAGhmtADRyu8A///8APbt9hT+8vbC//H2///x9f/+7/T/9+rv//Pl6f/q3eD/4NTZ/8rAxP+qoqb/mJCW/5ONkv+bl5r/pqOk/6qnqP+mpKX/oJ6f/5ubm/+XmJj/lJOV/5OSk8aanJ0UnJ6LAIqJ1wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC9uOgA///+AP3y+Rj/8/fL//D1///x9f/+7vT/9ujt/+/i5f/e0tj/mpKl/2BYbv9PS2n/UUtr/01HZ/9LRWH/VVJd/399f/+enZ3/op+g/6Cen/+cm5v/mJeY/5STlNuSkpYelJSTAGRjqwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD//+MA/fP4AP3z+Bz+8fbV/vD0///x9f/97fP/9uft/+rd4P+UjJP/VUx2/2xgsv93a9L/fnDa/31v2/9xZMb/SkaB/zg7UP99fIH/oJ6f/6Ogof+fnZ7/mpmZ/5STleqPjJ8tnJqPAGBcugAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA+/H1APrw9SL98fXf//D0///x9f/87fP/9+ju/7mvtv9WTnH/eGvI/5+S6v+7se//tarw/76z8f+soO//e2/i/0xKjP9QTlv/mpWa/6Ogof+fnp7/mpqa/5STlfGOjZdDl5h6AHx60AAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA/vP5APzy+Cf+8vbl//D1///x9f/67vL/7+Pm/312gv9mWq3/fnDm/7mw6v/59f//6+X9//v4///Wz/j/gHHx/3Nm0P9GQGD/iIOI/6WjpP+ioKH/np2d/5eWl/WTk5ZRlZVyAI+O3AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA/vP5AP3z+Sz+8vfo//D1///x9f/+8vb/3tTX/2dgdv92adf/eGny/7et7v/49f//6eL4//j0/f/Vzvj/f3Hw/35v7f9YTYX/dXJz/6elpv+lo6T/op+g/5qamviXmJlZoKGIAHl3zQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA//X7AP70+Sz+8vfp//D1///x9f//8/f/2c/T/2Zfdf96bNz/eGn3/7Cj8P/p4vv/5Nv4/+vj9v/Lwff/fm/z/4Bv9f9fUpT/c3Bz/6mnqP+mpKX/o6Gh/52cnPyYmZhpm5yWAHZ0xgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA//b7AP70+S3/8vfq//D1///x9v//8v"
    Sec08_HeaderLogo_2 = s
End Function

'--- Sec08_HeaderLogo_3 : Dashboard header, toolbar buttons (Layout/Update/KPI/Chart/Slicer/Save/Reset/Deploy) and the embedded company logo bitmap (chunk 3 of 18)
Private Function Sec08_HeaderLogo_3() As String
    Dim s As String
    s = s & "j/4Njf/2tke/96bNn/eWvy/7qt7//79///7ub1//fz+//W0Pn/f3Hy/39v7/9aT4z/d3R7/6uoqP+opab/pKGi/56cnf+VlZZ4np6JAIOBzQBoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA//b7AP70+S3+8ffq/vD3//7w9//98fj/9evx/4qAkfxxY7v/fW7q/7is7f/9+f//7+f4//by/v/Y0vj/g3Tq/4Bx2P9PSWv/joqP/66qq/+tqar/qKWm/5+env+ZmJuDqKSHAJmevgBoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA//X7AP7z+Sz+8ffp/vD3//3w9//98fj//vL4/8i7w/9nXYr/fHDP/6OX6v/IwPb/xLr2/8W89v+0qvT/gXLn/11Vlv9kY23/raet/6+rrP+tqar/q6ip/6Wio/+gnp+HxLCPAP//hABoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADv6eoA//T6AP3z+Sn+8fjm/vD3//3w9//+8Pj//vD4//fp7v+poKr/Zl2J/3VpwP9/cd//fnDl/39y4f97bs//YVeY/1xYaf+hnaD/trC2/7Svs/+xrK7/rqqr/6qmp/+in6GGzLiWAPn/iABoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADt5fIA//T5AP3z+h7+8ffX/vD3//7w+P/+8Pn///H4//7w9P/05+v/uq+1/3hxiP9nXI7/ZFmY/2NXkf9iWnn/eXZ9/6ypq/+9uLv/vbi7/7i0t/+1sLT/sa2v/66pqv+pp6d8ubSgAJ2jvwBoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADMxu0A///3APTr8xP98fXB/vD3//7w+P//8fn///L5///x9f/66/H/9+jt/+fa3v/Hu8H/saas/7Gnrf/Ct7z/z8bJ/8rDxv/Gv8L/wr2//7+6vf+6trn/tbG0/6+rrPuko6JlraueAImH3QBoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC4tOIA///0ANLM4QP+8vaZ/vH4//7w+f//8fr///H6//7u9f/76/H/+env//Tl6v/z5Of/7t/h/+fa3P/f1Nf/2c7R/9PJzP/Oxsn/ysPG/8bAwv/BvL7/vLi6/7eztvGoqKlGtbOiAJ+d3gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACamN0AXWHeAP///wDv4/Rk+u37+//y+///8vv///L7///x+P//8Pb//e30//jp8P/05ev/8eLl/+vd4P/l2Nz/4dXa/9vQ1P/Vy8//0snO/8zEyf/GwMP/vLe8/6ShsOGHiagkioupAP///A"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAf37TAL+o/wCekN1IvKry886+9//Yy/n/4NX6/+bb+f/n3Pj/6tz4/+na9//l1/P/3dLv/9bK6v/Sx+f/yr7n/7+z4P+zqNP/o5zE/5KLwf+AfLn/amew/1dUpsZGQ58WQD2YAImH1gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAeHbOAI1v/wCCbvMSi3L0t5J5+P+WgPX/m4bx/5+L7v+ei+z/m4no/5qJ6P+UhOj/joHl/4p84v+Dd9z/fHDe/3Zq2P9uY8//YVrJ/1lTzP9WUdL/UEzT+09Mzm8AAJMANyz/AHp5xwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAUEu1AFkr0QCTb/8AinP7T5F2/e2UeP7/lXn+/5Z7/v+Qef7/jHf9/4p2/f+Hdf7/g3L9/4Nw/v9/b/3/emz6/3Jl9v9qX+//Ylfo/1xU5P9YU9//U0/Wzk9NyB5CP8sAnZ7PAH17yQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX1u+AIuM1gCFZ/8AgnL5CI10/Z6Sdv3/lXr//5d8//+Sef//jnj+/4p3//+Jdv//"
    s = s & "hXP//4Rw//+DcP//e238/3Fl9f9pXu//Y1jq/1pT4/9XUt/5VVHXaHdn/wAjHrEAd3bEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wBjAP8AkXL/AIpy/C+PdP3ZlHj+/5d7//+We///lHj+/5F3/f+Md/7/hnT//4Ny//9/cP//eGv8/3Bk9v9oXe7/Ylfp/1lT4v9YVOG2XVrfElJO5wCpqtgAdHLAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABzcsAAv3D/AMcR/wCLc/1hknb985Z7//+We///lHj+/5N3/P+OeP7/h3X//4Nz//99cP//d2r8/25i9f9oXO7/Ylfp/1pU4+ZXVuJAWkz4AIBm/wCPjtwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABrabcAhIbOAIRj/wCFb/gIj3X7i5N3/f2Vef7/lHj9/5J3/P+NeP7/iXb//4Nz//98b/7/d2r8/25i9f9oXO7/Ylfp+1lV43piX+IDOzL0AJCP3ACOjdsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAjo3bABYA/wCIbP8Ah3T1F5F2/LOTd/z/k3f8/5N3/P+Pd/3/iXb//4Nz//97bv7/dWj7/21h9P9lWuz/YFfop1xX2xBRR+sA//+4AI6N2wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIKAzgCSev8Ajmz/AI1z+yqRdvzLk3f8/5N3/P6Pd/3/iXb//4Bw/v94a/z/cWX4/2pe8f9jWurKWFjiKFBQ+wBCN+4AiIfVAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wCJj8wAh3r/AJt6/wCNc/s3kXb8zpN3/P+Qdv7/iXT//35u/v92aPz/bWH1/2Va7thkWuk7UUX0AHFf1wBkYrAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74AOU9tAMuI/wCefP8AjXP7NpF2/MOPdP79inL//35u/v91Z/v/a17y1WZc7ElCG/8A9P/GAISD0ABoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAZ2W1ADNTSwCgdv8Annn/AI1z/CaNdP6GiHL+vH1s/cF5afuUcmPzMMj/EwAuHOsApKXZAI6N2wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA///gA///AAD//8AD//8AAP//wAP//wAA///AAf//AAD//wAAf/8AAP8GAAAgfwAA/gAAAAA/AAD8AAAAAD8AAPwAAAAAPwAA/AAAAAAfAAD8AAAAAB8AAPwAAAAAHwAA/AAAAAAfAAD8AAAAAB8AAPwAAAAAHwAA/AAAAAAfAAD8AAAAAD8AAP4AAAAAPwAA/gAAAAB/AAD/AAAAAH8AAP+AAAAB/wAA/8AAAAP/AAD/4AAAA/8AAP/gAAAD/wAA/+AAAAP/AAD/4AAAA/8AAP/gAAAD/wAA/+AAAAP/AAD/4AAAAf8AAP/gAAAB/wAA/+AAAAH/AAD/4AAAAf8AAP/gAAAB/wAA/+AAAAH/AAD/4AAAA/8AAP/gAAAD/wAA//AAAAP/AAD/8AAAA/8AAP/wAAAD/wAA//AAAAf/AAD/+AAAB/8AAP/8AAAP/wAA//wAAA//AAD//gAAH/8AAP//AAA//wAA//8AAH//AAD//4AAf/8AAP//wAD//wAAKAAAAEAAAACAAAAAAQAgAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjo3bAFRPuQBvSLIAAP//ADl92EA9ac+yLFOlxRtAdGsaNWsKAAAAAEZTfAA0NX8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wBJofgAAKr/ADuS6i44g9zNMnTL/ydcqf8bSIbwGj1xb1NGWwQAAAAAVlWnAF9bvgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF9bvgCGhNYAAIvwAD+K4AlAkeeYPIff/zN0y/8tZbj/IFCV/xw+dOQpOnM9AC0AAGNgxwBfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABYVLoASkOwADeg9QBBlu40QI/n4jmC2v8ydMv/LWW5/yFTmf4YQXv/Ijttlzg8bQWEd+gAW1a9AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAfHnPAH56zwA/ivwBP5TtdkSR6P46hN3/MnTL/y5luf8jVZz+GEF8/xo6b9krQGcnABgAAEFBjQBoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74AZWK3AJST6wBhRuEAYJrWB0eb8rFJm/L/QIzk/jN2zv8uZrn/I1Wc/hhBfP8aOGvwJz9pVx9BXQCUi+YAbGq4AGhmtABoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAZ2S2AGRhtwB2cNcAAAAAAHqd8Q9LnfXMSp72/0KR6f42fdX/LWa5/x9QlP4XOG/+FS9f/xg4bH0RO0YAAAUAANTB/wBWWaAAaGa0AGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAZWOxAGVjpgABHwAAa3AHABoofAg4To8hT4fLxE+Fz/8/dL/+Ml+x/yxTov4aQHn+HjJX/xsqSf8bK0OKMzhDJi1DbRMHAAAAAAAAAFBQYwB4dcQAAAAAAAAAAAAAAAAAAAAAAAAAAABSTMwAYF29AGJfuwBfW74AX1u+AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAaGa0AGlnswBjYLkAX1u+AF9bvgAAAAAAAAAAAAAAAAAAAAAASkmoABgYjQAAAAAALzhVGkxMUXRYWWCuUlZhx11fa+9UUmb/REZo/05Olf9ERZz/KC1X/ysqMv8nJi7+JCUt5ignLtQnKDK9LTRKYDQ5VQf///8AdHLA"
    s = s & "AAAAAABoZrQAAAAAAAAAAABCQo8AQUGNAGZhxgBiXroAXlq/AF9bvgBfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAYmCvAGBbwABlXe4AMknEAP///wDBqf4Aa2m3AGhmtADu6OsA7ujrAPDn8QATANoAgHMAAEhNXn5YWV35YGFj/2lpa/9kZWf/V1hb/0pJYP9dWaT/S0es/zEtVv8qKTH/JCMo/h8dIP8fHSD/IB4h/ycnLdIqLjMmMTcAAAAAAAAAAAAAZ2W1AEhHlQBDQo8ALjotAGF2CgAJAP8AcnGDAf///wAVHE4AYFy+AF9bvgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAammxAHJlzwCETv8AYlzjEThNuRgyQW4CAAAAAKWa/wBMS5sA8OvsAODb3ADc0tkAoJt9APbo/wtsam2bbmxv/3l3ev6DgYL/iYeJ/4OChP93dJb/b2e7/0pHrf9PTX7/YF1l/1VVWP5OTlH+SUVL/kI8Qv88OkHeVVBXTWFlfxccNK0Fgoo2AIN/kgBHR5QAQUGRAAAA/wA6MakEODenRzk3qns7OKwcMCe9AF5ZwABfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA////AG9sxwCshP8Aj3X5EG9i75hTVNjCRkTCUhwcdQNLQPsASkafAPHu4wClnJ4A5NvgIN3S13Hj19+lz8bJ3MzDyP/RyM3/zsXK/9PJzf"
    s = s & "/MxMj/vLTc/4Z81/9LSLH+enev/6Sfrf+WlJj/jYqN/4KAgf97eHj/cW1x/XJtcu5panLEaGZ2h313jiCIf3cAUk2HAEEf8wBJSqkGQkGyZT07tOQ+PbX+PjyxgTc2jgJuaNIAXFi5AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAH18ygB+fcoAeF//AHZo9URwYvXwZFnr/1lT3uZTT8d1QUCkF4Ru/wD///8A/fP4HfXq8LHw5On98eXp//Dj6v/w4ur/7eDq/+ve6P/g1N//1svb/8S64/+Jf93/Tkq6/29upv+YkqL/kIyX/4+NkP+Ihof/g4GD/4KAg/9/fIL/fXp//3l2eP92dHqbbm56DkBGQwA2NJAgNzediEBAru0/P7f/P0C1/0BAstQ/P6kkAACnAFFPqABoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACQj9sAkZDbAc2t/wJ6avaBal3w/2ha8P9lWev/YVfh+1NO0LJMSro6AABNAv/1+1n67/T1+Ozw//jt8P/26e3/8+Xr/+/i7P/t3+v/5tfo/93Q5//Nv+v/kYXg/1NPw/9iXqT/goCW/4aCjv+KiIv/ioiL/4qIi/+KiIv/iYeL/4iFiv+Fgof/enmI4l1jfzIzMpZIOjecwTg4pf07PK3/PT2x/j9Asv8/P7HwPTyqUjU1qgBCRngAc3OoAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA////AKKfzACYhfQHemr3tGld7/9nWfD/"
    s = s & "Zlnt/2Za7P9jWOf/YFjZ4llTzIvd0fKo++/1//jt8f/47fD/9unu//Pl6//v4uz/7N/r/+bX6P/dz+f/z8Hu/5SI5f9VUMT/Yl6p/4F/lv+Ggo7/ioiL/4qIi/+KiIv/ioiL/4qIi/+Jh4n/hYOI/oB8k/pRT4fDOTaW6Dk4pv84OKr/OTmq/zo5rP87O6/+PDqx/js5rH95fdwB7+DiAGVhuwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGZktgD///8AgXLtFHRm9M5pXe//aFvu/2hb7v9mWu3/ZVjt/2Va6f9tZN3/y8Dz//rv9f/57vL/+O3x//fq7v/05uz/8eTr/+3f6//m1+j/3c/n/8/C7v+YjOv/XFXP/2NgrP+Bf5b/hoOO/4qIi/+KiIv/ioiL/4qIi/+MiYz/iomK/4eFiP+CfpP/X1iT/zo4n/84OKr/ODiq/zk5qv84OKr/OTmr/zw6sv87ObCkOjakDJOcUABJSJoAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfXL4A87T/AIVy+yV1Z/TaaVzu/2hb7v9oW+7/aFvv/2ha7/9nW+z/dWvk/9nM+f/67/T/+e7y//nu8v/36+//9Ofs//Lk6//t3+v/5tfo/93P5//Pwe7/npDr/2Nb1v9kYa3/gH+Y/4eDkP+MiY3/jomO/4+Kj/+Pi4//jo2O/4uMiv+JiYj/hoOL/2djkv87OKD/ODip/zg4qf84OKn/ODip/zo6rP87Oq//OzmvrD03rhBWawwAQkKOAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYFy9AP///wCFdvwhdWf12Glc7/9oW+7/aFvu/2lb8P9oXe//Zlzn/3504v/i1/3/+u70//ru8//67vP/9+vv//Tn7P/y5Ov/7d/r/+bY6P/d0Of/0MLv/6CR7P9lXdn/ZWGu/39+nP+Gg5L/joqO/4+Mj/+PjI//j4yP/4+Nj/+NjIz/i4uK/4eFif9ua5H/PTyh/zg4qf84OKn/ODip/zk5qv86Oq3/Ozut/zo5rak8N6sPbn4jAEREkwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGlntwD///8AkYj8Dnpq+cVsXPP/alzx/2lb8f9pW/D/aFzv/2RY7P+Oger/6+D9//ru9f/77/T/+u7z//fr7//05+z/8uTr/+7h7P/p2+n/4dPn/9XE8v+gkuv/bGLe/2pmt/9+f6L/iIWU/4+Lj/+PjY//kI2Q/5COkP+QjZH/j4uR/4yLi/+IhY3/dHCY/kNCov84OKn/ODiq/zk5qv86Oq3/Ojqt/zw8qv86OamLLTV1A/7l/wBdWrUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABwbskAZWPFALeb9wN/bPmSbl/0/2xc8/9sXPP/a1zy/2pc8f9lWej/no7r//Pl/v/97vb//e/0//zu8//36+//9Ofs//Lk6//t4Ov/6tzq/+TU6P/WxfL/ppXu/3Zo4v9wa7v/gYCk/4yJl/+PjJD/j42Q/5COkf+RjpL/kY6S/5CMkv+Mi4v/iIaN/3t4nP9OS57/OTmm/zk5qv85Oav/Ojqt/zs6rv87"
    s = s & "ObDzODamWjAstAB4csoAa2m0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAi4rZAIyL2QBwUv8AhHP7TnNj9PNrW/L/a1vz/2tc8v9qXPH/ZFrl/6ue7v/57P7//e72//7v9P/97/T/+ezw//Xp7f/z5uz/7uHs/+rc6v/k1Oj/1sXy/6uZ8/96a+n/cWy+/4KBpP+Nipj/kY6S/5GOkv+Sj5P/k5CT/5OQk/+Sj5L/jIyM/4yLjf+Cf5//VlOh/z49o/86Oq3/Ojqt/zo6rf87Oq//Qz+3zj46rCD//wAAQ0SQAGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wBvbMgAlHf/AI52/RV8a/i0b2H1/2lc8f9qXfH/a1zy/2dc5f+2rfH/+vH9//3v9f/+7/T//u/0//ru8v/26u7/9Obs//Dj7P/q3Or/5NTo/9bF8v+tnPb/fW7s/3Jtvv+BgaT/jYqY/5OQlP+TkJT/lJKV/5WUlv+VlJb/lJOV/5CPkP+QjZD/h4Sd/2FdnP9APqP/Ojqt/zo6rf86Oq7/Pj2x+ENAtW9MSW0BTEqdAD09igAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAXlq+AF1XxACUfP8AhnL8RnRn9uhqYPH/a13x/2tc8/9sYeX/wbj1//vz+v/+7/T//u/0//7v9P/67fL/9uru//Tm6//w4+v/69zq/+TV6P/XyfD/r5/2/35u7P9zbr7/g4Gj/46KmP+UkZT/lJ"
    s = s & "GU/5WTlf+VlJf/lpWX/pWUlv+RkZL/kI2Q/4mHlf9taZj/Q0Gj/zs6r/46Oq7/Pz6y/0dFt7JGQrcVPju0AJuP/wBKSZYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF9bvgCytPEAgITAAIuA6wR+cPxocmT38Gla8f9qXPP/cGTq/8vD+P/89Pz//vD2///w9v/+7/b//Ozy//np7//y5ev/7uHs/+zd6v/o2ej/3s/t/7iq9P+Ecu3+eXG//4yGn/+TjZX/l5SX/5mVmf+Yl5j/l5aY/paVl/6VlJb/kZGS/5GOkv+Mi4z/d3WP/0VCov89O7P/Pz60/kpIur1TVMEqAAD3AFZRwQBfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74AYV2/AHd9twCDVv8AinL5CXlq+mdxYfXda1zz/3Nl7f/SyPn//vL7//7w9v/+8Pj//e72//vr9v/46PD/8OPr/+7h7P/r3Or/5tjo/97R6/+8r/P/gnPr/3xzwv+Ti6j/m5Gf/6Cbov+hm6L/oJyh/52bnf+bmJr/lpaX/5SSlP+UkZT/j4+O/4F/k/9LSaH/Pz648kRBuaFWVcYmAP8AADg4fABFRZMAX1u+AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACnp+oAhin/AKSD/wCPlfADhHH6NXNk96B7a+3s3s/4///y9//+8ff//vH3//7w"
    s = s & "9//77PL/9+ju//Lk6//t4Or/6tvo/+XY4//d0eD/y77o/62d5f+lncv/r6q4/66ps/+rqKv/qaao/6ekpv+joKP/n5yf/5uamv+Xlpj/lZSW/5GQkf+GhZj+WVasw09MxV1LSL4NAAC+AGVepgBIR5EAQkKPAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX1u+AF9bvgBiZZ8Al564AINq/wByXvYNqJfuVfHk+uT98fb//vL2//7y9v/98PX/+Ozw//Xo7f/z5ev/7+Hl/+rc5P/j193/28/Y/9XK0v/OxdD/xsDI/8C7vv+7trn/tLGz/7Crrf+ppqf/paOk/56cnf+bmpr/l5eY/5WUlv+QkJD/iomT929vq1V/dqQAIRn/AFhTvwBbWbEAU1OgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA////AJOS4gB9eM8At5T+APLn8Cv68fjc/vL2//7y9v/+8vb//O/0//fq7//05+z/8uTp/+3g4v/o297/4NTY/9rP0v/XzM//zsbJ/8bBw//BvL//u7a5/7Swsv+wq63/qaan/6WjpP+enJz/m5qa/5eYmP+VlZb/k5KU/4+OkPmOi5VjAAD/AIqJ1gB3dcwAX1u+AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAGhmtABuba4A29LzAAAAnQD68PY6/vL24//x9v//8fX///H1//7v9P/46u//9eft//Lj5//p3N//5Nfb/93S1v/OxMj/vrW5/7Gprv+ooKb/qqWo/6unqf+tqar/rqmr/6qnqP+mo6T/oJ6f/52cnP+YmJj/lpeX/5WUlv+RkJL8l5eXd////wCKidcAionXAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMbA7AD///8A/vP6QP/z9+b/8PX///H1///x9f/+7/T/9+nv//Pm6v/t3+L/5dje/87E0f+jmaP/b2hx/05KU/9IQ03/R0FM/0hEUP9dWWL/dnR4/5COj/+ioKH/pKKj/6Ogof+gnp//nZyc/5qZmf+Xlpf/k5KU/5SUlJQOJP8AAAAAAGxquAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACamtoA//77AP/090v/8vbt/vD0///x9f//8fX//e3z//bn7v/x5Oj/693f/83Cyf90bob/RT1f/1BIef9gWpb/aWGd/2Vemv9dVZH/SUBz/zMwTP9JR03/eXl7/5uamv+fnp7/oqCh/6Cen/+dnJz/m5ia/5WVlv+SkZSoiI27BLjDvgBpZ7MAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA6+fkAPzy+QD98vhV/vH18/7v9P//8fX///H1//zs8v/25+3/8uTp/9jN0P9va3H/S0Rq/3Fkr/98btb/em3e/4Bx5f+Bcuf/gXHn/3hs1f9bVZ//LTJV/z9AUP+GhIn/oJ2e/6Ogof+in6D/np2d/5qamv+Wlpb/kY+WvYmD0AwAAAAAX1u6AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOLZ9QD88vQA+vD0YP7w9fn+8PT///H1///x9f/87PL/9ufu/+7h5v+NhJD/SEFg/3hqwP+Ke+b/pprq/6yh6P+lmen/raHs/6yg6f+eker/fW/m/2Bctf8xM1T/XFph/5qVm/+joKL/op+g/56dnf+ampr/lpaX/5GQlMqPjKEd1tM2AHVyzAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADf1vYA//z5AP3y92X+8fb7//D1///x9f/98PT/+ezx//fp7v/KwcT/VU5i/2Zbqf99ceL/i3/f/+HZ+f/28P//4dr+//Lu///49P//zML3/31u7/99buj/VU2L/0E8T/+PiI//pKGi/6Ogof+gnp//nZyc/5mYmf"
    s = s & "+SkpPYkZCYLJWVOQCOjdsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA8+r1AP/5/AD+8/lu/vL2/v/w9f//8fX///H1//ru8v/26e3/oZqd/1JLcf95a9v/fG3u/4+C3//n4vj/+/r//+rj+//18Pz//Pv//87G9v99b+3/emvy/3Nowv88NlH/d3J2/6WjpP+lo6T/pKGi/6Cen/+cm5v/lpWW2ZeYmS6lqEUAjYzbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7ujrAP/4/AD/9v8A/vP5cv7y9v7/8PX///H1///x9f/98vX/8ebq/4F6f/9eV4b/fm/u/3Nm9f+Lfeb/4tz6//r4///m3vf/8Or5//v4///Oxvb/fnDt/3lq9P+Ec+T/TkNz/2ViY/+npKT/paOk/6akpf+joKH/npyd/5iYmOCZm5o3rrGFAHBtxwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wDh3+EA//n/AP7z+XP+8vb///D1///x9f//8f"
    Sec08_HeaderLogo_3 = s
End Function

'--- Sec08_HeaderLogo_4 : Dashboard header, toolbar buttons (Layout/Update/KPI/Chart/Slicer/Save/Reset/Deploy) and the embedded company logo bitmap (chunk 4 of 18)
Private Function Sec08_HeaderLogo_4() As String
    Dim s As String
    s = s & "X///L2/+/k6f95cnj/Zl2L/35v7v91Z/r/invr/9HI+P/n3/j/4Nj8/+fd9v/o3/j/wrf1/31u8P95avj/h3Pr/1NHfP9kYWL/qqeo/6ajpP+mpKX/pKGi/6Cenv+ZmpnpmZqYQZyflgBlYrYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA4t/hAP/6/wD/9Pl0//H2///w9P//8fX///H1//7x9//06PD/f3h+/2Vdi/9/cO7/eWv5/4x96P/k2/n/9fH9/+zj9f/u6PT/9fH9/83E+f99bvP/eWr7/4Ny5v9SR3//ZWJp/6qmp/+npKX/pqSl/6Shov+gnp//m5qa9JOTlVCUlYwAh4bRAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7ujrAOLf4QD/+v8A//T5dP7x9//+8PX//vD3//7w9v/98Pf/9+/3/5WOlf5hWof/fnDm/3lq8v+MfeT/6uH6//38///v6fb/7+j3//z7///Qyfb/fnDp/3xt8/9/cdP/Rj5q/29sdf+rp6j/q6ip/6qnqP+mo6T/oJ+f/5uamveRkZZXl5eBAH9+yQBoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wDi3+EA//r/AP/0+XT+8ff//u/2//7w+P/98Pf//fH5//3z+f/KwMX8YlR1/Xlqy/9+b+v/inzh/+Te+v/9+v//8Oj5/+rk+v/8+v//0cr3/4Bx5v+Ec+z/dmmy/0I+VP+OipD/rqqr/62pqv+tqar/qqeo/6Kgof+dm5z/mZecYqGdjAAXCKsAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA4uDiAP/5/wD+8/ly/vH3//7w9v/+8Pj//e/2//3x+f/98fj/8OTp/4Z7i/9iWJP/gnbY/4t94f/LwvX/3db7/9nQ+//Px/r/3Nb7/7uw9v9+b+r/fXDX/0pFcP9kY2r/rqiu/6+qrP+tqar/ramq/6uoqf+opab/op+g/qCeoGS0p5YAUUm8AGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7ujrALO/vQD/9/8A/vP5cv7w9/7+8Pb//vD4//3w9//+8Pj//vD4//3u9f/Vyc7/aWN3/2Vdmf+EeNb/hXfl/4Z36f+Eduj/hnjo/4Z55/+HeOb/fW7X/1NMgf9JSVL/nJmc/7Kstf+0rrP/sa2v/6+qrP+tqar/rKip/6Shov6fnZ9ktKeWAFFJvABoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wDx5/QA//n7AP7z+Wf+8fj7/vD2//7w+P/+8Pj//vD4//7w+P/87vT/+u3x/8rBxv9vZnn/XVWB/3BmtP9+cdP+em7b/3xv2/97bcv/bGKp/09Ib/9cWGf/mZWY/7eytf+4s7f/trG1/7Sws/+yrbD/sKut/66pq/+ppab9oZ+hYrarlwBNRbsAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//jwAP/1+gD+8/lT/vH28f7w9v/+8Pj//vD4//7w+f//8fn///L2//zu8v/46vD/3NDU/5ePlf9mYHn/W096/1tPg/9XTYP/WUx1/1xVZ/98eXz/rKqr/725vP+9uLv/v7q9/7u3uf+4tLf/tbG0/7KusP+wqq3/ramp+KqpqFazsaAAlpfTAGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANTM8AD///8A+vD1Pv3x9uX98Pb//vD4//7w+P//8fr///H6///y9//+7/T/+erw//bo7f/w4+f/28/V/7mutP+kmqH/n5ae/62kqv/Ct7z/0MbJ/8vDxv/GwMP/xL7B/8K8v//Au77/vbi7/7i1t/+1sbT/"
    s = s & "sKyu/6qoqOqen51CoKGVAI6N3QBoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACsqeEA////APHm7SD+8fXU/vH3//7w+P/+8Pj///H6///x+v//8Pj//e3z//nq8P/46e7/9ebr//Pk6f/y4+b/7d7g/+fa3P/j19n/3dHV/9jN0P/Tycz/zsbJ/8zFyP/GwMP/xL7B/8G8v/++ubz/urW4/7WwtP+vrK7cmZqbMaGlWQCRj9sAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAfX3YAEhStADx5/gF/vP4qP3x+P/+8Pn///H6///x+v/+8vv//u/4//3s8//87PL/++vx//bn7f/y4+j/8uLm/+3f4f/p2t3/49fa/9/U1//d0dT/2M3Q/9LJzP/Oxsn/zsbJ/8rDxv/GwMP/wr2//7+6vf+9uLv/t7S5vaGipBSJjXQAw8DkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI+O2wCPjtwBBA+HAezg9HL57P39//L7///y+///8vv///L7///x+v/+8Pf///D2//3u9P/66vL/9eju//"
    s = s & "Tl6//y4ub/7d/h/+nb3v/k19z/49bb/97S1//YztL/1cvO/9PKzv/Px8z/ycLH/8jCxf/BusD/sKy0/4uMp6UtNowDXWGfAISD2QAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACBf9MAgX/TAP///wCdkNRjwbHy/drK+P/h0/r/6dz9/+3j/v/y6P7/9On+//Tp/f/36f7/9uf9//Xm+v/w4/j/69/z/+ba8P/j1+7/4dXu/9zP7v/Sxej/zcDl/8K32P+4sc7/q6PK/5uVw/+Nibf/c3Cq/15bnv9JRZmQcG/LAgAAAACOjdsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAXFi8ADozpwCabP8Ah3P0KY118NGUe/P/mobx/qCQ7/+pmO7/r5/s/7Gj5/+xo+f/sqPl/7Kj5f+yo+X/q5/l/6ec4/+fk93/mo/X/5WM0v+Mg9b/hXrR/35yyP91ar3/ZmKx/1lVuP5WUbz/UEzB/01JyP9QTMzoU1DOORwZwQB+fcsAgH7MAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF9bvgBm"
    s = s & "YsIAG0gtAGx10gKKcfZ5kHX8+pR5/f+Vevv/lHr3/5V79v+Wffb/k3z2/4x48f+LePD/iXbx/4Ny8v+AcO7/g3Lv/39v8P96bOz/d2nt/3Zo7P9wZOj/bGHj/2Rb4P9dVtz/X1fe/1tV3/9VUNv/UE7Qm0xMtwoAAIYAjYzaAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAZmPCAIiA2gCWYv8Ah3b7H410/cCTd/z/lHj9/5V6/v+We///lnv//454/v+NeP7/i3f//4p2//+Jdv//hXP//4Rx//+DcP//gXD//3tu/f92aPj/bWL0/2hd7/9jWOr/XVTl/1pT4/9WUd3/Uk/V5U5Mx0AeHqsAZWaWAGBdqwBoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACQj90AjIrXAKZt/wCLc/xbj3X99JN3/f+Vef//l3z//5Z7//+PeP7/jXj+/4t3//+Jdv//iHb//4Rz//+Ecf//hHD//4Nw//97bv3/d2n4/2xh8/9pXu//ZFnq/11U5f9ZU+L/V1Le/1RR1pVQTb0HXWCCAD4/gABoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjo3bAHFwxwCeff8AhnH7EIxz/qaSdv3/lHj+/5Z6//+XfP//lXr//5N4/f+QeP3/jXf+/4p2//+FdP//hHL//4Nx//+AcP//emz9/3Vo+P9sYfP/aF3u/2RY6v9cVOT/WFPh/1lU49hcVt8tV0//ADMzhQBDQ5AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB2c8kAf3rWAJBj/wCJcvkyjHP+15J2/f+Wev//l3z//5d7//+Vef//lHf9/5F3/f+NeP7/h3X//4Rz//+Bcv//fHD//3hr/f9yZvn/bGHy/2db7f9jWOr/XFTk/1hU4fZVVt5uX3aoAv///wB8e8kAQkKPAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAARUKQAJqZ5wCIjsQAbc9wAYpz/mKRdv3xlXn+/5d8//+We///lXn+/5N3/P+Rd/3/jXj+/4h1//+Fc///gXL//3xw//94a/3/cWX4/2tg8v9nW+3/Y1jq/1xU5P9aVeWsWFriEgUi/wAAAAAAjo3bAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGhmtABqaLYAfIS2AIBN/wCCbfkKjnT7h5J2/PmUeP3+lXn//5R4/f+Td/z/kXf9/414/v+Jdv//hnT//4Fy//97b/7/eGv9/3Fl+P9rYPH/aFvu/2RY6v9bVeTbVVfhOyg/4wCTlN8AjYzaAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFJQngBcQ9IAjmz/AIR08xeOdfu0k3b8/5R3/f6Td/z/k3f8/5J3/P+OeP7/iXb//4Z0//+Bcv//em7+/3dp/P9xZPj/al/x/2Za7P9hV+nxXVblYQvACQD/ygoAfXvJAI6N2wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAi4rYAJWW3ADzBf8AjnT5LpJ2/MuTd/z/k3f8/5N3/P+Td/z/j3f9/4l2//+Fc///gHD//3ls/f91aPv/cGT3/2pe8P9iWO"
    s = s & "n9XlfnjV9Yzgpsb3IA////AMnL/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wCSk9gAo6PdALB5/wCLcvxGkXb835N3/P+Td/z+knf8/454/v+Jdv//hHH//3tt/f93afz/cGT4/2xg9P9nXO7/YlrpsU9Z3BspNecAAAAAAGNisACOjdsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjo3bAHRzyQCnhP8AgnDvAo1z+1KRdvzdk3f8/5N3/P+Odf7/iXX//4Rx//96bP3/dmj8/25i9v9pXfH/YljsxGhe6ClMTfYAVla8AFpXsABCQo8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74Ay5X/"
    s = s & "AP///wDpqP8AjXP7UpF2/N6Td/z/j3T+/oty//6EcP//eWz9/3Vn+/9sX/T/Zlrv0WJa6DVHKf8AlpbYAI2N1wAuH/UAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGBcvQBaVcQA/8H/AHdm6AKMc/xEkHX9w45z//mLcv//gm/+/3hr/f92Z/v1a13xtWhf6jr/+sgBkJLaAJST2wCOjdsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74AX1y9AIlz7wDckP8A+VH/AI1y/SSNdf5riHL+nYBu/ax8a/2eemr5ZIVv+RdvUv8Am6HSAJOW1wCOjdsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD////AA////////8AB////////gAH///////+AAf///////4AA///////+AAA///"
    s = s & "////wAAB//////+AAAHwf//+B4AAAGA///wAAAAAAB//+AAAAAAAH//4AAAAAAAf//gAAAAAAA//+AAAAAAAD//4AAAAAAAP//gAAAAAAA//+AAAAAAAD//4AAAAAAAP//gAAAAAAA//+AAAAAAAD//4AAAAAAAP//gAAAAAAB///AAAAAAAH//8AAAAAAA///wAAAAAAD///wAAAAAAf///AAAAAAD////AAAAAAf///8AAAAAD////8AAAAAP////wAAAAA/////AAAAAD////8AAAAAP////wAAAAA////+AAAAAD////4AAAAAP////gAAAAA////+AAAAAD////4AAAAAH////gAAAAAf///+AAAAAB////4AAAAAH////gAAAAAf////AAAAAB////8AAAAAH////wAAAAA/////AAAAAD////8AAAAAP////wAAAAA/////AAAAAD////8AAAAAP////4AAAAA/////wAAAAH/////AAAAA/////+AAAAD/////4AAAAf/////gAAAD//////gAAAP/////+AAAB//////8AAAH//////4AAA///////wAAH///////gAA///////+AAH///ygAAABIAAAAkAAAAAEAIAAAAAAAAFEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACOjdsAf3zRAAD//wAwoe4DPHjVRz9m0awuU6nAH0J1cRo3axL///8AXVm9AGBdvABoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wCOjdsAjY/cAQD//wA4jeg9NoHa1zJzyv8nW6j/GkiG9Rc+dYkuPGwJpJb/AFdXnwBgX64AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wCNjNoA/6PiAEeN3xI/kee3OYPc/zJzyv8tZbf/IlSc/xtAePApPWlgViv/AVpZmQA/QIgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI2M2wCNlOEABKX/AD6U7Fg/jeXwOYLa/zJzyv8uZrv/IlSc/xhEgf8eOnG/Nz10HwQdJgBdXMAAU1KgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGhlxAAA//8AMJHwDUKU66NAjOT/N3/Y/zJzyv8uZrr/JVii/hpGhP8aOm/0IDpqWgAzXQB3asoAWlmmAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0ALS29gAQpv8APJnzMEOW7dRFkur/OYPc/zJzyv8uZrr/JVii/hpGhP4YOXH/Iz1smkJIYQpbRckBR0eUAGhmtAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGdltQBmZLYAZ2S6AGBbvQApoP8ATpztU0qd9PVKm/L9P4ri/jR40P8uZ7v/JVii/hpGhP4XNmv+HTlpyyZBcRkHM1QAY2KvAGVjsQBoZrQAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AGhmtABmY7UAUE2wAOjW/wAAhdcAUpz1YEqe9v1KnfX8QI3m/jV71P8tZrn/IVKZ/xc8df4VLVz/FjNk6hk5bycYNDgA////AHx5xgBnZbMAaGa0AGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAaGa0AGZksACjlv8AHigAAAARcwMdGHkPSXi0ZVKGzPNNgMj+Omy2/jJbrf8tUqL+GkJ9/h81Xv4eK0v/GClE5SUuOkUxPVEpKUR4Dh8xKwB2c8kAYWCoAGdlswAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74AX1u+AGNgugBfW74AX1u+AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AGhmtABpaLMAYl+6AF9bvgBfW74AAAAAAAAAAAAAAAAAAAAAAAAAAABta7YATkyoALWi/wAmMlsNQUJIWVhYXZpQU1+2VVlk11xcbPtOS2P/QEVr/09Nl/9GRZ3/KC9j/ykpNP8qKC7/IyUv9ycnLdcoKTLOKCk1rS86Vk00NmAEvKz/AJ2c6gBbWacAaGa0AAAAAAAAAAAAAAAAAEJCjwBFRZQAdW7hAGJfuQBgXL0AX1u+AF9bvgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAaGa0AGhmtABpaLMAYl+6AEpFsgBzccsAjo3bAAAAAAAAAAAA7ujrAO7o6wDm4OgA///3AAAAAABCSmFoVVdd7F5fYv9naGn/aGhq/2BiY/9TU1f/Sklk/2Baqf9NSa//NDFk/yopM/8oJy7/IR4i/h8dIP8fHSD/IB8j/ykp"
    s = s & "McgwMjcpAAAAAHp5yAFaWokAZGG5AF9bvgAAAAAAaGa0AFNSoAA8PIoA8df/ACIhlgAAAH4AVVG7AGFdvABoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD0x5ABnZLUAa2e6AAAAAAAAQlEAAEgpAMbG2AHIyv8AWVufAEREkQBMTJUA6ubnAOTf4ACpmaAAvbK7ADkzAABZV1yKXFxf/mRjZ/9wb3D+c3N1/3JydP9mZ2z/XlyD/2Zgs/9JRq//Pjx2/0ZDTf9BQET+OTk8/jc2Ov4zLzX+MCsx/y0sM9QxMjYwAAAAAI6q/wD//wAAOza8AFpYzgBBQY4APDybAHh2uAAOGMsAMzWYEzY1okk8OK4gFAqyAGVhuQBpZ7MAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGVjtwBkYbsAABgAAIlz8hFgXOR9P03BbTs5phZDRsEAZmK0ADEzfACCgLcA6+jjAOzi3QDQysoI3dLcKObZ6VWuqamzo56h/rGrr/+xrK7/ubO2/8C5u/+3sbj/qqHT/4B20v9MSLL+ZGOk/5aQo/+OjJH/goGE/3p5e/9xbXD/amZo/2FcY/NvZ27BaWl0hk5WeEt2bKASTlBFAAYBWgAADSIATUecASIgmABCP7QhPDmynDw6su08OrGaODifEAsRggDWwf8AX1u+AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAF9axwBqaKMA////AH9r+WpsYPH1XFXi9k5L0KRGRLUnAAsAAUlHlgGEgrsA//n7APzy9hvq3+WL5Nnd5ujc4ffn2+H76Nvi/+nc5v/l2d//4NTd/9rP1f/OxdH/vbPk/4F21f9MSLX/amms/6OdsP+cl6H/lZKW/46Ljv+Eg4P/gX9//3t4ev91c3j/dHJ3/HFucvR0cXqfdXaCFioqsAFpYLEBDTQvAzg4njVCQq+3Pj62/T4/uP9AP7TvPj2vUFJJ2gCBeOAAYl7DAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIaF1AD//wAAg3X5EHFk9KtsXvP/ZVnt/19X5f5aU9HDSke6VSQshQgAAJkA////BPzx9nj36/D49uvv//fr7//05+3/8+Xr//Di7P/u4ez/6dvo/+DT4//WyuL/xLjp/4l+3/9RTcH/XVuk/4qGnf+MhpX/jImQ/4yKjP+Ihoj/h4SH/4eFif+Fgof/hYGH/4B9gP98en71cnF7YVhRiQA8ObIOOTSVZzg3oNI+P67+P0C1/j9Atf4/QLL/QUGxlkREpwlQTo0BS0qdAGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AMrJ/wB+U/8AjXX9MnNl89doW+//Z1nw/2VZ7P9kWej/WVHc7VRQyY1NSrgozcj2Ef3y+L357vL/+Ozw//js8P/26e7/8+Xr//Hj6//u4Oz/6tvq/+TV6P/azef/ybvu/42B3/9VUMX/Wler/3x6mv+FgY//iIWM"
    s = s & "/4qIi/+KiIv/ioiL/4qIi/+KiIv/iYeK/4aDiP6Cfor/bnCMoSwydzQ4NpmjOjef8zg4qP85Oqv/Ozuu/j4+sf4+P7L+PDusxjw6qhUuLK4AiofMAGtpsQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AGZksQB2Xv8Agm/6TXJk9PRnW+7/Z1nw/2Za7P9mWuz/ZFnq/2Rb4v1bVdDSgnjZlfbp+On67vP/+O3x//jt8f/26e7/8+Xr//Hj6//u4Oz/6tvp/+PU6P/azOb/zL7x/5CE5P9WUcb/W1et/3x7m/+FgY//iIaM/4qIi/+KiIv/ioiL/4qIi/+KiIv/iYiK/4aEh/+FgY//cGyS7EA8iuE6N57+ODip/zk5qv85Oav/OTmr/zo6rv48OrL+Ozqv6zs7qjIAAJQAY1+7AGNfuwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAZ2S2AFdbmABbQv8Aemv0Z29j8v9oW+7/aFvu/2hb7v9mWu3/Zlnu/2Va6/9kXOL/kofk/+7j+v/67/P/+e3y//jt8f/36+//9Ofs//Lk6//v4ez/6tvp/+PU6P/azOb/zL7x/5OH6f9dVtH/Xlqy/3x7mv+FgY//iYaN/4qIi/+KiIv/ioiL/4qIi/+MiY3/i4qL/4iGiP+Fgo7/dnCW/0xGk/84N6b/ODip/zg4qv85Oar/ODiq/zk5qv87Oa//PDqz9jo4rFVNTH8BR0aYAUdGmAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX1y+AEBJiwGZf/wDfWz4fW9i8v9oW+7/aFvu/2hb7v9oW+7/Z1rw/2hc7v9mW+T/oJXu//To+//67vL/+e7y//ru8//37O//9ejt//Pl6//v4uz/6tvp/+PU6P/azOb/zL7x/5uM6f9jW9j/X1u0/3x6m/+EgZL/i4eO/4yJjP+Pio//j4qP/4+Lj/+PjY//jIyL/4qKif+Hhoj/fnqQ/1BMlP84N6b/ODip/zg4qf84OKn/ODip/zk5rP86Oq7/Ozqv9jw4r11GQ4MBQkKOAUJCjgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX1u+AERIkgGaivwDfW74fG9i8v9oW+7/aFvu/2hb7v9pW/D/aFzu/2he7P9nXtz/raLx//fr+//67vL/+e7y//ru8//37O//9ejt//Pl6//v4uz/6tvq/+PU6P/azOb/zL7x/5yN6v9mXtv/X1u0/3t6nf+CgZf/jIeP/46Kjv+PjI//j4yP/4+Mj/+PjY//jY2M/4uMiv+Ih4j/gX6N/1ZUlf84OKb/ODip/zg4qf84OKn/ODip/zo6rP87Oq7/Ozqu9js4rlxHRIUBQkKQAUJCkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAZWK4AFBPlABePv8Af3L8ZHJj9v9qW/H/aVvv/2ha8P9pW/D/aFzu/2Za7v9sYOT/v7T1//jr/P/67vP/++/0//ru8//37O"
    Sec08_HeaderLogo_4 = s
End Function

'--- Sec08_HeaderLogo_5 : Dashboard header, toolbar buttons (Layout/Update/KPI/Chart/Slicer/Save/Reset/Deploy) and the embedded company logo bitmap (chunk 5 of 18)
Private Function Sec08_HeaderLogo_5() As String
    Dim s As String
    s = s & "//9ejt//Pl6//v4uz/697r/+bX6P/dz+j/0MD0/5uN6f9qYN//ZF+8/3p6pP+DgZj/jIiP/4+Nj/+PjY//j42P/4+NkP+PjZD/j4yQ/42LjP+JiIr/g3+R/11amv86Oqb/ODip/zg4qf84OKr/Ojqs/zo6rf88PKr/PDur8Tc3pULUz2oAVlOtAVZTrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAdXamAP///wCDZ/8AjXf9QXZl9+dsXPP/bFzz/2xc8/9rXPL/aV3w/2ZZ7v91Z+X/zsD3//rs/P/87/T//O/0//vu8//37O//9ejt//Pl6//v4uz/7N7r/+ja6f/h0en/0sH1/6CP6/9zZ+L/bGbC/3x8qP+HhZv/joqS/4+Nj/+PjY//kI6R/5GOkv+RjpL/j4yU/42Ljf+KiIv/hICT/2hln/9BP6D/ODiq/zg4qv85Oar/Ojqt/zo6rf87Oq//Ojir1zc3nR0aF6MAami5AGpouQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AIOB1ACXYP8Ainb9Hnlp98BtXvP/bFzz/2xc8/9sXPL/al3x/2VZ7P98b+P/3Mz5//zt/P/+7/T//u/0//zu8//37O//9ejt//Pl6//v4uz/7N7r/+ja6f/h0en/0cH1/6WS7v95auj/bmjE/319qP+Jh5z/j4yT/5COkf+QjpH/kY6S/5GOkv+RjpL/kY6R/42Mjf+KiYr/hoWS/3NwpP9FQZv/Ozun/zk5rP86Oqz/Ojqt/zo6rf88O7L/PDmvnzYzngtbXn4BSkmX"
    s = s & "AGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAA///bAYVx+3pwYfP6alzy/2tb8v9rXPH/a1zx/2RY6f+DeOX/5dr6//zv/P/+7/T//u/0//7v9P/67fL/9uru//Xn7f/x4+z/7d/r/+ja6f/h0en/0cH1/6qY9P99be3/cWrG/318qP+Jh5z/kI2U/5KPkv+Sj5P/k5GU/5SRlP+UkpT/lJGU/4+Oj/+NjYz/jYqV/3l2pP9LSZz/Pj2m/zo6rf86Oq3/Ojqt/zs6rv9CP7XqSkW6TXdh/wBdXKUAVlWjAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGBcvwBdWMEAq4P/AIdw/CV9bPnRbWH0/2ld8P9qXvD/bFzz/2Za6/+OhOb/6eP9//3x+f/+7/T//u/0//7v9P/67vL/9+vu//Tn7P/y5ev/7eDr/+ja6f/h0en/0cH1/6uZ9f9+bu7/cWrG/318qP+Jh5z/ko+V/5SRlP+UkZT/lZOW/5WUlv+VlJb/lZSW/5KRk/+QjpD/jYqV/4F+oP9XUZr/Pj2m/zo6rf86Oq3/Ojqu/z49sf5APrSaPDauDQAA/wBNS6QAKSp3AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF9bvgBfW74AYlzDAAAA/wCBcPpgc2b28mtg8f9qXPL/bFzz/2hd6/+WjOj/7ef+//3x9f"
    s = s & "/+7/T//u/0//7v9P/77vL/9+vv//Tn7P/y5Ov/7eDr/+ja6f/i0+j/1Mjy/62d9v+Ab+//c2vH/359p/+Lh5v/ko+V/5SSlP+UkpT/lZSW/5aUl/+XlZj+lZSW/5KSk/+QjpH/jouP/4SDl/9hXJb/Pz6o/js6r/46Oq7+Pz6x/0lGt9JMSLorLimuAF1XzABcWLkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74AAACAAK/fXgCGevUJf3H9enJj9/RpWvH/alzy/2ld7f+glu7/8u3///7y+P//8PX///D2//7v9v/87PL/+urx//Tm7P/w4uz/7d/r/+vc6v/m2Oj/28zv/7ao9P+Fcu/+eG7K/4aBov+Ui5n/lZGV/5mVmf+Zlpn+mJeY/5iWmP6Wlpf+lZSW/5KSk/+Rj5L/j42Q/4qJi/9qZ5P/Pzym/z08s/8/PbT+Ski60VFTwUC638IAXVnAAF9bvgBfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74A////AKGb+wCRRP8Ah2/6EXlq+nhxYfXoa1zz/2lb7/+nm/L/9ez///7x+P/+7/b//vD4//3u9v/76vb/+ury//Lk6//v4ez/7d/r/+nb6f/l1+j/3M/t/7ms8/+Bcez/eG7K/4qFq/+YjZ//nZaf/6Cbov+gmqL/n5yg/5yanP+bl5r/l5aX/5STlP+TkZT/kZCR/42Ojf9zcZX/QkGk/z4+uPlDQLi3VVLGOElhsgM3Mm0B"
    s = s & "SkibAF9bvgBfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjo3bAHRxywBoTO8ArYX/AImD+QeCcPlFd2b4sWxe7/KxofP/++77//7w9//+8Pf//vD4//7v9v/77PP/+enw//Pl6//v4uv/7N3p/+ja5//j1uP/3dDj/8a57P+kkuj/m5DP/6ehuP+rpLT/qaWs/6ilp/+npKb/pqOl/6Ofov+fnJ//m5qa/5iXmP+VlJb/k5KT/5GQkf97epz7UE2v0k9MxXdIRbsX//+9AHZuqgB5cakAQkKPAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF9bvgBfW74AZ1zZAJ63nwBmQv8AgGr7FHpp7F/RwfTN/O/4//7x9f/+8vb//vL2//3w9v/47PD/9unt//Tm7P/x4un/7d/l/+fa4v/i193/2c3Z/9TJ1v/Nw9b/w73O/8G8wf+8uLr/t7O2/7Gur/+uqav/qKWm/6WipP+fnZ7/m5qa/5mZmf+VlZb/k5OU/42Ojf+EgpvjVFW6P7mlxQE9Ob0AZWLCADAuiABcXKkAQkKPAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaWe1AKal/QBEQK8A///KAf"
    s = s & "/16Qj06vqY/PL4//7y9v/+8vb//vL2//zw9P/46/D/9ejt//Pl6//w4uX/7N7i/+bZ3v/g1Nn/2c7S/9bLzf/Oxsn/yMHE/8K9v/+8uLv/t7O2/7KusP+vqqz/p6Wm/6WjpP+enJz/m5qa/5mZmf+VlZb/lJOV/5CQkP+NipLigH6dN///AAB1dcEAZ2S/AF9bvgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AG5trQCsp+AAfH3kAerh7BH67vSp//L2//7y9v/+8vb//vH1//zv9P/46/D/9eft//Pl6//u4eT/6dzf/+TX2//b0NP/2M3R/9bMz//Mxcj/wry//8G7vv+8t7n/tLCy/7Gtr/+vqqz/qaan/6WjpP+fnZ3/nJub/5mZmf+Wl5f/lZWW/5KSk/+Qj5Dql5eZS6apFwCQj90AkI/dAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADg1/QAx7/wAvvx+hT+9Pix//L2///x9v//8fX///H1//7v9P/46/D/9ejt//Lk6P/r3eD/49bb/97S2P/TyMz/sKmu/5ePlP+KgYn/g3qD/4Z/hv+PjI/+npud/6ejpP+qpqj/qaan/6ajpP+ioKH/np2d/5uamv+ZmZn/lpeX/5STlf+SkZLymZqZXqSmgQCC"
    s = s & "gc8AgoHPAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACvrOQAAAAAAPzx+hb/9Pi6/vH1///x9f//8fX///H1//7u9P/36O7/8+br/+3g4v/n2t7/18za/6igsf9uZXH/SkZT/zw6SP89OUj/PTlJ/zk1Rf9AO03/UUxX/2dmav+Jh4j/n52e/6Ogof+ioKH/oZ+g/56dnf+cm5v/mZmZ/5aVl/+UkpT8k5OUd5maewCJh9sAT0yaAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADf2ucA//n5AP/09hr/8/bG/vD1///w9f//8fX///H1//3t8//25+7/8uXp/+3f4f/d0dT/koqZ/0hBYv9GPGj/X1aV/25msP92bLb/cmmz/25kr/9fVJf/PDZs/zAuQP9IRk3/d3h6/5uamv+enZ3/oqCh/6KfoP+enZ3/nZqb/5mXmf+UlJX/kZCUibyCAAD//7UAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAADu6OsA/fL5APzz+h7+8vfS/vD0//7v9P//8fX///D1//zr8v/25+7/8uTp/+jc3/+XkpX/RUBT/11Tif98bsb/e27c/3hr3v9/cOf/f3Dn/4Fy6v9+cOP/c2jJ/0tKhP8lLEn/R0dW/4qHjP+fnZ7/o6Ch/6Ogof+fnp7/nZyc/5iYmP+VlJX/kI2ZnoB3/wYAAHoAX1u7AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA+/H0APjv8yP88PTf/vD0//7w9P//8fX///D1//zr8v/35+7/8+br/7iutv9LRlf/XlOH/4Z13f+Nf+b/pJjp/6WZ5f+fkub/ppnq/6WY5/6jluf/iHro/3hr3f9MTZX/LjFH/2JhZv+alZv/o6Ci/6Ogof+fnp7/nZyc/5iYmP+VlJb/kJCUrpOPqBL//wAAcW7JAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA//f4APrw9if+8vbi/vD1///x9f//8fX//fD0//rs8f/36O7/49fc/3Jrdv9OQ3P/eW/O/3tu4v+ckOP/6+P+//Ps///c1P3/7ej///Xw///o4P3/mo3v/3pq8v95a9X/RD5r/0VBUP+QiZH/pKCi/6Ogof+hn5//np2d/5ua"
    s = s & "mv+Xlpf/j4+Swo6MmiCPjAAAjo3bAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA//b6APzy+Sv+8/jq/vD1///x9f//8fX//vD0//nt8P/36u//yL7B/1NNXP9mWqf/fm/r/35w5f+hlt//9O/+//r4///o4f3/8+79//38///t6P3/m4/r/3Zn8/97ber/ZFyj/zYxRv98dXv/paKj/6WjpP+koaL/oZ+f/56dnf+amZn/lZOVxZiamyKrsQAAj47bAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wDu6OsA//T6AP3z+DL+8/js/vD1///x9f//8fX///H1//zw8//37O//oZqd/01IY/97btD/eGn4/3pr7/+gk+P/8/D+//v7///p4fj/7uf1//38///t6P7/nZLr/3do8/97bO//f3DS/0A4W/9iYGH/pKKi/6Wio/+mpKX/pKGi/6Cen/+bmpr/l5eYyZmamyW5vlQAf33SAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAO7o6wDu6OsA//b9AP3z+DL+8/js/vD1///x9f//8fX///L1//3y9f/26+//joeM/1VPaf+ActX/cmb4/3do8v+bjej/5+H+//Lt/v/h2fz/6OD5//Tv/v/n4P3/m4/s/3do8/97a/T/iHTm/01Ccf9cWlv/qKWl/6WjpP+mpKX/paOk/6KfoP+dnJz/mJmY2Jqcmi+prJEAbGnAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wDu6OsA//j+AP70+TP/8/jv/vD1///x9f//8fX///H1//7x9v/47PH/joiM/1dQav+Bc9b/dWj6/3xt+f+Zi+f/3dT8/+Xc9f/k2/v/5dvz/+bd8v/Z0Pr/mInw/3do+P97bPj/hXLi/0xCcv9dWl3/qaan/6akpf+mpKX/paOk/6Ogof+enZ3/mpqZ4ZaYljeXmZQAk5LjAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wDu6OsA//j+AP70+TP/8/fv/vD0///x9f//8fX//vH1//3w9v/68Pn/lo+W/1dQa/+Bc9b/eWv5/3pr8f+gkef/+PH+//r4///x6fb/7efy//v5///w7P//npLw/3Zn+P97a/j/g3Lb/0xCd/9gXWb/qKSl"
    s = s & "/6ilpv+npab/paOk/6Ogof+fnp7/mZiY75CQlUSUlIcAhYTOAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wDu6OsA//j+AP70+DP/8vjv/u/1//7w9//+8Pf//fD2//3w9v/78/v/ta2z/llSbf55bcj/fGz0/3pr6/+fkOX/9+/+//z7///x6vf/6eD0//v6///u6v7/m5Hm/3pq7/9/b+3/e27D/z01Wv9tanT/q6eo/6yoqf+sqKn/qqeo/6Sio/+fnp7/mZiZ8pCQl0iXl4IAfHvGAGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wDu6OsA//j+AP70+DP/8vnv/u/2//7w+P/+8Pj//fD2//3y+v/98vn/5Nne/W5ic/trXKH/gXHo/3xt5v+cj+T/8uz+//v5///y6vr/5t74//n3///w6///nZLo/3xs6/+JeOf/amCd/0A8T/+PipD/rqqr/62pqv+tqar/rKip/6elpv+gn5//nJuc/JmXnVKinY4AV1KzAGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wDu6OsA//f+AP7z+DL+8vnt/u/2//7w+P/98Pf//e/2//3y+v/88fj/+Ozx/6OYoP9YTXn/e3DF/4J14f+ZjOP/3dX6/+Tf/f/i2v3/1Mv8/+Lc/f/b0/3/lIfs/35u7P96bsv/Qj1h/2Vkav+tp63/r6ut/62pqv+tqar/rKip/6qnqP+lo6T/oZ+f+p+eoFSzp5cAX1u3AGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wDu6OsA//b9AP3z+DL+8vns/u/2//7w+P/98Pf//fD3//7w+P/+8Pj//u/2/+fZ3f99doP/V1B+/31zxf+HeuH/iXvn/4p86f+Kfej/in3n/4x/5/+Mfuf/hXbp/3psz/9KRm//RkZP/5uYm/+xqrX/s62z/7Ktr/+vqqz/rqmr/62pqv+rp6j/op+g+p+eoFSyppcAX1q3AGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wDu6OsA//b6AP3z+C3+8vno/vD2//7w9//98Pf//vD4//7w+P/+8Pj//O71//rr7//c0tb/dG18/1tQef9uZa3/fnHV/4N14/99b+b/fnDn/4Fy4f9+cdL/b2Sv/0xFcf9STV3/"
    s = s & "kY6R/7axtP+2sLX/trC1/7WwtP+yrbD/sKut/66pq/+sqKn/paKj+p+doFSyppcAX1q3AGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8A/fP6APzz+iH+8vjY/vD1//7w9//+8Pj//vD4//7w+P/+8Pj///L3//zv8v/46/D/2tDV/5GIjv9bVWr/WFGA/2RZmv9nXar/Zlyr/2VYm/9VTnn/TEhb/2dmav+gnaD/ubW4/7m1uP+9uLv/ubW4/7aytv+1sLT/sq6w/7Gsrv+vqqz/rKmo9Kelp0u3s50AUUywAGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADm3vMA////AP70+Bj+8/fA/fD1//7w+P/+8Pj//vD4///x+v//8fr///P3//7w9P/76/H/+Orw//Hj5//Hv8P/mI+b/3dqff9qXnT/ZFpz/2xhdv+Bd4L/oZqe/7y3uf/DvsH/wry//8G8v//CvL//v7q9/7y4uv+4tLf/tLG0/7KusP+uqqv/q6in6qampD6pqKAAeXbJAGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADBu+sAeXniAe3k7hH68PSp/vH2//7w+P/+8Pj//vD4///x+v//8fr///L4///w9f/87PL/9ejt//Tm6//15+v/8ePn/+HU1//Sx8n/zcPF/9HHyf/XzM//2s/T/9XKzv/Mxcj/ycLF/8bAw//DvcD/wry//7+6vf+8uLr/ubW3/7axtP+xra//qaen1JeZmC2gonsAjYzaAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC8ueEAAAKtAKGmxAH88PV8/vL2//7w+P/+8Pj//vD5///x+v//8fr///H5//3t8//76/H/+Onv//rp8P/z5On/8eLn//Hi5v/u3uH/6dvd/+TX2v/f09b/2c7S/9bMz//Tycz/z8fK/87Gyf/Jw8X/xsDD/8S+wf/BvL//vrm7/7m1uP+2sbX/rautwpWWlyCgphkAlpTcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADiCOAP///wD/9PpO/fL57P3w+v/+8fn///H6///x+v/+8vv//vD5//3t9P/97fP//ezz//rq8P/05uz/8+Po//Pi5v/v4OL/6tze/+bZ3P/h"
    s = s & "1dn/4NTX/9zQ1P/YzdD/0cjL/87Gyf/Px8r/zMTH/8jCxP/DvsD/wLu+/8C7vf+9ubz/trW7lZyenAd3fHYA5N7pAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAkI/cAP///wDi1esq8+b6z/zv/v//8vr///L8///x+v//8vv///H6//7w+P//8fb//vD1//3s9P/56vH/9eft//Pl6//y4ub/7uDj/+vc3v/m2d3/49fc/+LW2//d0dX/2M3R/9XLzv/Tys//0MfM/8rDyP/HwsX/ycHF/7y3vP+npa//fX+hguHfwwBXW5gAiYjaAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAg4HUAJ1/OACGfqsjr6HnytbF+P/i0/r/6Nr7/+7h/v/x5v7/9er///fs///37P//+Oz///nq/v/46f7/9+j5//Lk+f/u4vP/6d3x/+fa7//k2O7/4tXu/9zO7v/UyOf/0MTm/8i82v/AuNP/tq3P/6igy/+bl8D/hoKv/21qpP9WVJb5Qz+XcwAAOwA+QnwAj47cAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADAOLAD9bZQCHd/oIi3XwjpJ67/yahPL+oZDv/6mb8P+xo+//uavv/72x6f+9sen/vrHp/8Gy6f/Bsun/wbLp/7qu6P+2q+j/raPj/6ec3f+lnNX/n5bY/5WM2f+NgtH/hXrK/3pwuv9saKv/XVmx/1dTtP9QTbb/Ske9/k1KxP9QTcrTVlLOJAAArQCDgc8Ag4LQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAbmvHAHJtzQD/dv8AhnDzL4xx9uCSdvv/k3n7/5N79/+UffL/lX3x/5Z/7/+Wf+//kn3s/4x75/+Me+f/iXjo/4N06f+AcuT/g3Pk/4Bw5f97beD/dmng/3Vn4v90ZuD/bmLa/21i2P9mXNf/XVbV/15X2P9fWNv/V1Ld/1NP1/xOTMqDXVvHA4aFzwGop/UAU1GfAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOjOnAJiY3QCCh8UBhXbsBYpz/ICQdf32lXj+/pV6/v+Vef7/lXr+/5d8//+Tev//jXj+/414/v+Ld/7/inb+/4h1"
    s = s & "//+Ec/7/hHL+/4Nw//9/bv7/fG79/3lr+f91Z/f/bWL0/2hd7v9iV+n/XFTk/1tU5P9ZU+H/VlHd/1JP1dhQTsspGBbOAGtpuQBvbbsAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX1u+AGFfrQBWUpsAhDz/AIV0/CaOdfzGk3b8/5N3/P+Vef7/l3z//5d8//+Qef7/jXj+/414/v+Kdv//iXb//4l2//+Ec///hHH//4Rw//+EcP//f3D//3pt+/90Z/f/a2Dx/2ld7/9lWev/X1Xm/1tT4/9XUuD/U0/X/FFO0IBHR5oF///5AH17yQBpZ7UAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGhmtABsarYAlP8AAJ2L/wCKc/1jjnT+8pN3/f+Vef//lnv//5d8//+Vev//j3j+/454/f+Ld///iXb//4h2//+Ec///hHL//4Rw//+EcP//fm7+/3ps+/90Z/f/a2Dx/2le7/9lWev/XVTl/1hT4f9YU+H/WFPfxlZTzyVtaP8APj6GAD47iQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABcWLwARju4AAByAACIcfwRjHP9qZF1/f+Ud/3+lXr//5d8//+XfP//lXn//5R4/v+ReP3/jnf+/4t3//+GdP//hHP//4Ny//+Bcf//fG7+/3dq+/9zZvf/a2Dx/2dc7f9kWOr/XVTl/1hT4f9YU+H0W1fjYAwF/wCLi9cAkZDZAGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74AeXfIAIB8zwCTYf8AiHP4Motz/teRdf3/lXn//5d8//+XfP//lnr//5R4/v+Td/z/kHf9/414/v+Idf//hHP//4Nz//99cf//e27+/3Zo/P9xZfj/a2Dx/2db7f9kWOr/XVTl/1hT4v5VVuCnT1zOD4qAwAGdneUAjo3bAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAyAKGg7gCQnMEBfaHCAYpz/mGRdf3vlXn+/5d7//+XfP//lXn/"
    s = s & "/5R4/f+Td/z/kHf9/414/v+Idv//hXT//4Nz//99cf//e27+/3Zo/P9vY/b/al/w/2db7f9kWOr/XVTl/1pU5NVcV+c1STn6AGJfqwB0csMAjo3bAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AJqa6AC1mv8AfCz/AIJt+g2OdPuGknb8+JR4/f6Vev7/lXn+/5N3/P+Td/z/kHf9/4x4/v+Jdv//hnT//4Nz//98cP7/em39/3Zp/P9vY/b/al/w/2hb7v9kWev/XVXl81VW4G5oWukF//+mAHp4yABnZbUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABvbbsAZVfGAJVs/wCFdPIVjXX7spN2/P+Td/3+lHj9/5N3/P+Td/z/kXf9/414/v+Jdv//hnT//4Nz//98b/7/eWz9/3Vo+/9vY/b/al7w/2da7P9hV+n+XVbmmk1b0wsLHv8AOTO1AFdTqwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAkI/dAI6N2wHVJf8Ai3T2MJJ2/MyTd/z/k3f8/5N3/P+Td/z/k3f8/454/f+Jdv//hnT//4Ny//98b/7/eGr9/3Rn+v9vY/b/aVzv/2RY6v9gV+m/X1fbIf//AACSkuMAjo3cAJihfwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjo3bAI6O2wCYpcYAOMJnAIxy/EiRdvzek3f8/5N3/P+Td/z9k3f8/454/v+Jdv//hnP//35u/v96bP3/dGf7/3Fl+P9sYPP/aF3u/2BY59xVV+FGAAC+ADg6bABjYrAAjo3bAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    Sec08_HeaderLogo_5 = s
End Function

'--- Sec08_HeaderLogo_6 : Dashboard header, toolbar buttons (Layout/Update/KPI/Chart/Slicer/Save/Reset/Deploy) and the embedded company logo bitmap (chunk 6 of 18)
Private Function Sec08_HeaderLogo_6() As String
    Dim s As String
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wBcdJEAAFwAAGVc6AKNc/tdkXb85JN3/P+Td/z/kXf9/413/v+Jdv//hXL//3xt/f94a/3/c2b6/21i9f9pXfH/ZFnr6mhd6loAZIUBxapqAAAACgBCQo8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wDLzv8AZGqlAKt5/wCHdPQDjXP7Y5F2/OSTd/z/k3f8/490/v6LdP/+hnH//3tt/f93af3/cWX4/2te8/9lWe7vYFfqcHNz0gWBfcsBRkKzAGBcvgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AG1nwgD/1/8Af2r1A41z+16Rdvzfknb8/490/v6Lcv/+hnH//3tt/f93af3/cGP3/2ld8fNmW+yAcWrgCv//YQCWleAAfXvQAF9bvgAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AGlnswBvZssA5I3/AGhf6AKMc/xKj3T9v41z//aMc///g3D+/3lr/f93afz8dGP3zmVa7GZ6cOUN//+rAIiH3gCPjtsAjo3bAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGhmtABoZrQAaGa0ACZxAACAgN4CjXL+Iox2/l2Ic/6Ogm79n35s/Zp7a/ptgG34Lv///wBkdJoAlJLhAI6N2wCOjdsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD////8AB////8AAAD////4AB////8AAAD////4AB////8AAAD////4AA////8AAAD////4AA////8AAA"
    s = s & "D////wAAf///8AAAD////AAAH///8AAAD///+AAAD///8AAAD///8AAAD8H/8AAAD/+B8AAAA4D/8AAAD/8AwAAAAQB/8AAAD/4AAAAAAAB/8AAAD/4AAAAAAAA/8AAAD/4AAAAAAAA/8AAAD/4AAAAAAAAf8AAAD/wAAAAAAAAf8AAAD/wAAAAAAAAf8AAAD/wAAAAAAAAf8AAAD/wAAAAAAAAf8AAAD/wAAAAAAAAf8AAAD/wAAAAAAAAf8AAAD/wAAAAAAAAf8AAAD/wAAAAAAAAf8AAAD/4AAAAAAAA/8AAAD/4AAAAAAAA/8AAAD/4AAAAAAAB/8AAAD/8AAAAAAAB/8AAAD/8AAAAAAAB/8AAAD/+AAAAAAAH/8AAAD//AAAAAAAH/8AAAD//wAAAAAAf/8AAAD//wAAAAAA//8AAAD//8AAAAAA//8AAAD//8AAAAAA//8AAAD//8AAAAAA//8AAAD//8AAAAAA//8AAAD//8AAAAAA//8AAAD//8AAAAAA//8AAAD//8AAAAAA//8AAAD//4AAAAAA//8AAAD//4AAAAAA//8AAAD//4AAAAAA//8AAAD//4AAAAAA//8AAAD//4AAAAAAf/8AAAD//4AAAAAAf/8AAAD//4AAAAAAf/8AAAD//4AAAAAAf/8AAAD//4AAAAAAf/8AAAD//8AAAAAAf/8AAAD//8AAAAAAf/8AAAD//8AAAAAA//8AAAD//8AAAAAA//8AAAD//8AAAAAA//8AAAD//+AAAAAA//8AAAD//+AAAAAA//8AAAD//+AAAAAA//8AAAD//+AAAAAA//8AAAD//+AAAAAA//8AAAD//+AAAAAB//8AAAD///AAAAAD//8AAAD///gAAAAD//8AAAD///gAAAAH//8AAAD///wAAAAH//8AAAD///wAAAAP//8AAAD///8A"
    s = s & "AAAf//8AAAD///8AAAAf//8AAAD///+AAAA///8AAAD////AAAB///8AAAD////AAAD///8AAAD////wAAD///8AAAD////wAAH///8AAAD////4AAP///8AAAAoAAAAYAAAAMAAAAABACAAAAAAAACQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACOjdsAhYPVAF5avgBrYa4BOYTfCkpmz0lJXdedM1G4tixGdHMhPXMfKTZrA2BcwAFfW74AY1+6AGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACOjdsAkZDdAGBdvwE1jusHNoriaTR70+kycMj7KV6t/B"
    s = s & "pIhvYWP3bDGjhqLUJCoAFgXL4Bfny7AHFuvAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wCOjdsAjo3bArGBxAE4k+xLOobg5TV81P8yccf/K2Cx/yJTmf4ZR4X7Fz1zvjM+aBsAD00AWFelAVpYqwBfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAI6N2wCOjdsAAHHRAE2H2hRAlOfCPYni/TV81P8yccj/L2i8/ylerf8fTpD+Gzxy8yw+ZpE2MbEGRESQAm1n1QBfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX1u+AIF/0wCOjtsCAKf/ADmR6WNBj+bqQYzj/zV81P8yccj/L2i9/yhdq/8aSYn/GESB/xo5cds0OoI7FiZXAF5bvQFgXb0AaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAWVW6AFlVugAnt/8ANYnpC0SY77M+i+T7Nn/X/zR70/8yccj/L2i9/ylerf8hUpb9GESB/xs6bvMqPmaRDBVkAmtkzQGGg8YAeHXDAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAcW7JAHFuyQIFtf8ANZXxQEOR6N5Ekej/PIfh/zV81P8yccj/L2i9/ylerf8hUpb9GESB/xg5cP8ZOW/QKj1kIgAqLwBRUJ0BUVCdAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAhoXTAJCP3AEyo/IAPpryeUWV7e1Glez+P4vj/zV81P8yccj/L2i9/ylerf8hUpb9GESB/xg5cP8cO2/iM0RlWixDVwBBQY4CSEeUAGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABdW6kAXlupALXBsgD/lAACSJzyrUqd9PpLnfT8RZLp/jiA2f8zdcz/L2i9/ylerf8hUpf9GESC/xg4b/4aNmjxJ0BqiAVLRAAABUAAeHbEAG5rugBoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAaWizAGJeuwB3dckAgYDOAQ8A0gCjjsEOTp3zyEud9P9LnfT8RZLp/jyI4f81fNT/L2m+/ylerf8eTY7+GD95/hYxY/4VL17/Fj11si8wYgq1iv8AT0+cAnZzwQBoZrQAaGa0AGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGhmtABoZrQAaWiyAGFeuwFWU7sA////ACAnjADrnv8FRpvyukme9v1Ln/j8RJTt/jmA2v8xdMv/LGW3/yVZpP4YRYP/Fzdt/RQqWP8VLVr/FzJcthMzrAMpNTkAADsZAAAAAABjYq4BaGa0AGhmtABoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AGhmtABkY7AAd3CrABEXAABISjYAAABeAxohqBEnOHEmUoC6tVaAxPpOeb3+SXK2/ilYoP43Vaj/M0+l/xxFgv0YOW3/IzRV/iEtSf8cKEH9Fyg9tzMzNDYxO0o4L0t8HwA2/wAmOVkAbm2yAWRjrwBoZrQAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AGlntQBsa7AAGisAACQmLiNLSU5eVVddiUxPYZ9GTVy1XV9t5FtYbv1PSmX/RkRl/zdDdP5STJz/SUae/ys2ff4hKEj/LCox/yspL/8mJSz9Iicz6CopLscqKzTFKSw5ryovQno1RnIrJTgAAH93tA"
    s = s & "B2c8IAdXLAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX1u+AF9bvgBgXL0AZWK3AF9bvgBfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AGhmtABoZrQAZ2W1AGBcvQBfW74AX1u+AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsAqKTOADY1lgANIjcALD1oMUNJVqFYWVvnXmBi/2FiZP9iY2X/ZWVn/1xeX/9WVFn/SUZV/1BPe/9lXLD/UU2u/kE9jv8rJz//LCs0/ysqMv8mJCr/IyAk/yIgJf8iICT/IiAk/yYmL/ksMD+cMDI5IQkQAABERJEBYWCuAF1bqQBoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAABCQo8ASEiZAP/z/wBgXL0BZWK3AV9bvgFfW74AX1u+AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAaGa0AGhmtAFoZrQAZ2W1AV9cvQEAAGcAhYTWAI6N2wAAAAAAAAAAAAAAAAAAAAAA7ujrAO7o6wDu6OsA7efrAP///wAAAAAATFBhalRXXu9cXWD/XF1g/2pqbP5pamz+Zmdp/15fYv5UV1r/RURV/01Nev9mYLf/"
    s = s & "R0W0/z47jf8pJj7/KSgw/ycmLf8fHSH+HRse/R0bHv0dHB79HRwe/SEfI/8qKS7MMDM2PBsfAABgX60DhoTTAV1boQBjYLoAX1u+AAAAAAAAAAAAAAAAAGhmtABlY7EAOzyJAEVEkQIZHYcAJyabAAAAgQBeWr4BZGG5AGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGhmtABoZrQAaGa0AUxJnABqXu0AOFDEABIljQCUkuACl5bkAGhmtABoZrQAAAAAAAAAAADu6OsA7ujrAO7o6wDu6OsB5d/nANjJ8wB+eHEAVlNYb1dXWvRfXWL/ZWNn/m9ubv50c3T/eHd6/3Z2d/9vb3D/YWF5/2hjnf9oYbn/Ske0/kJAkv9CQWT/UU1T/klKTf9CQkX+Pj9C/j09Qv45NDz+NS81/jAtNf8rLDPHJiotNVpWVwAbMCoAAAA7AFhXsQFeWr8AYV28AGhmtAAAAAAAQkKPADEyfwBoZrUBhISzACs6ZgIxNJUsNTSgazw3rzIYG4oAZWSyAWJeuwBfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGlnswBpZ7MAZ2O2AHuRmgNqXuxnOVHBei06mi1NVIMCfny9AWlntQE2NoMAQkKPAO7o6wDu6OsA7+jsAOTg4QDk298ASDhpAO"
    s = s & "jZ+xzy5Ps9qKSklZKPkPShnKD/qaSn/6qlp/+tqKr/u7S3/7u0t/+2r7P/q6bB/6CV3P94bs//UEy4/klIof99e6X/lpCe/4yKkP+CgYT/e3p9/3Z1d/9va2//aWRm/2JeYv9dWGDtbWRsvndveYpMV3BaNUSOJVdF/wOGfp8Ab2ysAGhmtAA3OIQAQkKPAEJCjwJpaZAAMyinAzw3sE08ObHGOziw6jw5s7AvNJEb////AF9bvgFfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AGJfugBgXb0BkXP/AIpy+llsYO/bWlfh71BO074+OrRNMSmiAnBptABCQo8CRkaTAK2p0ADu6OsALYEeAP///wTg2Ns12dDTld3R1srj193W3tPX6dfN0Pze09v/39Tb/9rQ0//XzNT/2c7U/9fMz//Oxcn/w7zX/7Gl7v9zasf/T0u3/kpJpv+Rjrf/qKK0/6CcpP+XlZj+ko+S/4yJi/+Eg4P/gH5+/3x5ev9zcHL/cG1x/3BucvVpa3HgaWZxznhxhHKJhKQRjol/AHt5vAConf8AQkKPAlxQqwFdXMcESEe2Zj8+s9c9O7X/Pj63/z8+s/FAPbN5PDqsAFxZvABfW74AX1u+AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAZ2WzAGdlsgFpaKAAUE3VCXpo+bBu"
    s = s & "YvT8aFvu/1hT4P1UUNrYSki6aDU8ggxDRIEBQ0ORA4KAvwHg5dQA8+n0A/rx9ETz6O7F7ODn+PDk5//w5Oj/7uLp/+/h6f/v4ev/7eDq/+zf6P/r3un/39Pd/9vP2f/SyNj/wrvY/7er7v9yaNL/VFDD/0lIpf+Fhaz/oZuo/5mQoP+Rjpf/k5CT/42LjP+Fg4T/gX+A/4J/gv+AfoD+fHp//315gf95dnn/dnNz/3NxdO1xcnhlcnSACiwrmQJWU68CNEdaASIybBk2NZ+HQUGn4z9At/8+P7j9PkC4/UBAs/9AP7LFLzOSHQwVdQBgXL8BYFy/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAi4rYAIuK2AJwTv8AgnX7M21i8thsXvP/aFrw/2VZ7P9hV+f/X1jW51FMw6w2OJwtNjeSAdDJ6wHEvfAC/fP6G/zx95357fL/9urv//jt8f/37O//9ent//Pm7P/z5ev/7+Ls/+7h7P/t3uv/5tjn/+HT5v/azeX/zsLl/7yu8P99ctj/V1LI/0ZFrv9xbpr/h4Wc/4mDkv+IhJD/jIqN/4uJi/+Jh4r/iIaJ/4mHiv+Jh4r/hoOI/4eEif+Fgof/g4CE/oF/hP92dHy5bWt3I4SBbwBDP64CPDesSjgzj7o4OKbwPj+u/kBAtP9AQLT/P0C0/j9Asv9BQbPjR0azWEdHyABFRZUCRUWVAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAj47bAI+O2wKGYf8AinL8a21h8epoXO//Z1nw/2dZ8P9kWOr/Zlrq/1xT4PpRTdHRSkm3ZTxAigz///8A/PL8Mvzx+Nb67vP/+Ozw//js8P/47PD/9urt//Tm7P/z5ev/8OLs/+7h7P/s3ev/59np/+PT6P/bz+b/0cPo/7yu8f9/dNj/WFPI/09Ltv5pZaD/gH+a/4WBj/+Gg47/ioiL/4qIi/+KiIv/ioiL/4qIi/+KiIv/ioiL/4qIi/+HhYn+hYKI/oN/h/9zdIzoaW6ONBYoXhU1NJ2GPDiY3Tk3o/w4OKn/OTmr/zs7rv88PLD+QECy/j5Asv4+PrDwODejfzY1ogBBQY0AKix3AGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQA////ANLO5gFpK/8AiHP9nm9i8fhoW+7/Z1nw/2dZ8P9kWOr/aFvu/2Za7P9iWOb/ZFzb61NNxbNST8xN4dj4VP3y+PP67vP/+O3x//jt8f/47fH/9urv//Tm7P/z5ev/8OLs/+7h7P/s3ev/59no/+PT6P/bzub/0cPo/8O1+P+Addn/WFPJ/1FNt/5oZaP/gICa/4WBj/+Gg47/ioiL/4qIi/+KiIv/ioiL/4qIi/+KiIv/ioiL/4qIi/+Ih4j/hoOH/oWBjP9/e5f+TVR3qjc0hcM7N5XxOTik/zg4qf84OKn/ODip/zg4qv86Oq3/PDuv/Tw7sf49OrL9OzmsrGtj9wg5L/8AZ2W1AWdltQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAZ2W0AYiFnACOmqIEeWn5tGxh8f5oXO7/Z1nw/2hb7v9nW+3/ZFnq/2RY6v9jWOn/ZVrn/2Rc2/phWs3trJ7l7Pzu+/767vP/+O3x//jt8f/47PD/9uru//Tm7P/z5ev/8OPs/+7h7P/s3ev/59no/+PT6P/bzub/0cPo/8S2+f+EeOH/WVTK/1FNuP5oZaP/gICa/4WBj/+Gg47/ioiL/4qIi/+KiIv/ioiL/4qIi/+KiIv/ioiL/4qIi/+Ih4j/hoSH/oWBjf+BfJj/ZVqT+EI8kvw5OKP/ODip/zg4qf86Oqz/Ojqs/zo6rP84OKr/Ojqt/js6sP48OrP/OjqsyTM2mSMSGYMAXVm5AV1ZuQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABnZbUAZ2W1AY97/wCBdOUYeGn2y2xh8f9oXO7/aFvu/2hb7v9oW+7/aFvu/2dZ7/9mWe//ZFnq/2dd6f9zad7/wrb1//bt+//57vL/+u7z//jt8f/47fH/9+zv//bp7f/05uz/8+Xr/+/i7P/s3ev/59no/+PT6P/bzub/0cPo/8S2+f+GeuP/Y1vZ/1hTxP9oZaP/gICa/4WBj/+Gg47/ioiL/4qIi/+KiIv/ioiL/4qIi/+KiIv/jYmN/42LjP+Kior/iIeH/4aDi/+BfJf+aWSR/0Q/k/84OKb/ODip/zg4qf84OKn/ODip/zg4qf84OK"
    s = s & "n/ODip/zk5rP89OrX/PDqz1js3rUAzKswAQ0ORAkNDkQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74AX1u+AamC/wCKcv0peWr30mxh8P9oW+7/aFvu/2hb7v9oW+7/aFvu/2dZ8P9nWfD/aFvu/2Za6f91a+D/yLv3//vw+//57vL/+u7z//nu8v/67vP/9+zv//bp7f/05uz/8+Xr//Di7P/s3ev/59no/+PT6P/bzub/0cPo/8W3+f+PgeL/ZV3a/1lUxv9oZaP/gICa/4WBkP+Hg4//jYmO/4uJjP+Pio//j4qP/4+Kj/+Pi4//j42P/42MjP+LjIr/iYiI/4eGhv+Ggo3/cm2T/0ZBlP84N6X/ODip/zg4qf84OKn/ODip/zg4qf84OKr/Ojqt/zo6rf87Oq//Ojqu1zw3r0I1K9IAQkKPA0JCjwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74AX1u+AZuF/wCGd/01eWr31mte8P9oW+7/aFvu/2hb7v9oW+7/aFvu/2hb7v9oXO7/aV/t/2de4P96ceL/1sn+//vv+v/47fD/+u7z//nu8v/67vP/9+zv//bp7f/05uz/8+Xr//Di7P/s3ev/59no/+PT6P/bzub/0cPo/8W3+f+QguP/amHg/1lUxv9oZaP/gICc/4GAmf+Hg4//joqO/4yJjP+Pio//j4qP/4+Kj/+PjI//j42P"
    s = s & "/42Njf+LjIr/i4yK/4iIh/+Gg4z/eXaR/0lFlP84OKb/ODip/zg4qf84OKn/ODip/zg4qf84OKr/Ojqt/zo6rf87Oq//Ojqt1zw3r0I2K9IAQkKPA0JCjwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABgXL0AYFy9AbmY/wCLefsgeWr3z2xh8f9pXO//aFvu/2hb7v9oW+7/a1vy/2hb7v9pXu//aF3s/2Rb3P+Ge+H/3dL///vv+v/57vL/+u7z//ru8//67vP/9+zv//bp7f/05uz/8+Xr//Di7P/s3uv/59no/+PU6P/bz+b/0sTp/8W3+v+QguP/a2Lg/1tVyP9oZaP/f4Ce/4GAmf+Hg4//j4qP/4+Lj/+PjY//j42P/4+Nj/+PjY//j42P/4+Nj/+NjIz/i4yK/4mIiP+Gg4j+eneQ/09Olv86Oqb/ODip/zg4qf84OKn/ODip/zg4qv85Oar/Ojqt/zo6rf87O63/Ojqt1js2rT4yKckARESTAkREkwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAaGa0Aaen/wCPkvcOfW/9xG9g9f9sXPP/aVvv/2lb7/9oWvD/aFrw/2hc7v9oW+7/Zlnv/2db5/+Wien/6N////nr+v/57vL/++/0//vv9P/67vP/9+zv//bp7f/05uz/8+Xr//Di7P/u4ez/6dvp/+bY6P/d0Of/2cju/8a3+v"
    s = s & "+QguP/bWPh/2Nc1f9qZqf/f4Co/4GAmv+IhJD/j4uP/4+Nj/+PjY//j42P/4+NkP+PjZD/j42P/4+NkP+Pi5D/jIuM/4qJiv+Gg47/gHuY/lNQmf87O6X/ODip/zg4qv84OKn/ODip/zo6rf86Oq3/Ojqt/zs7q/8/Pqf/OjmtxjA0kSALFHUAX1u8AV9bvAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAWVS7AAAAkQD//wAAiXT9m3Fh9fdsXPP/bFzz/2xc8/9sXPP/bFzz/2le7/9pXvD/Zlnu/2pe5v+kk+3/7+P///nr+//87/T//O/0//zv9P/67vP/9+zv//bp7f/05uz/8+Xr//Di7P/t4Oz/7N3q/+fZ6P/j0+j/28ju/8a3+v+Tg+T/dmnj/2df1v9xbbD/f4Co/4SDnP+Oipb/j4uP/4+Nj/+PjY//kI2Q/5GOkv+RjpL/kY6S/4+Mlf+PipL/jIuM/4qJiv+Gg47/gHyY/l5aof9DQaD/OTmo/zk5rP84OKn/ODip/zo6rf86Oq3/Ojqt/zs6r/46Oq78Ojippx88CgP///8AaGa0AWhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAb2zIAG9syAKWd/8AinL7Z3Nl9upsXfP/bFzz/2xc8/9sXP"
    Sec08_HeaderLogo_6 = s
End Function

'--- Sec08_HeaderLogo_7 : Dashboard header, toolbar buttons (Layout/Update/KPI/Chart/Slicer/Save/Reset/Deploy) and the embedded company logo bitmap (chunk 7 of 18)
Private Function Sec08_HeaderLogo_7() As String
    Dim s As String
    s = s & "P/bFzz/2lb8P9rXPL/Zlns/21j3f+xnu7/9uf///rr+//+7/T//u/0//7v9P/77vP/9+zv//bp7f/05uz/8+Xr//Di7P/s3uv/69zq/+fZ6P/j0+j/28ju/8a3+v+ciOj/fG3p/25k1v9xbrD/gICo/4iFnf+Oipb/kIyQ/4+NkP+PjZD/kI6R/5GOkv+RjpL/kY6S/5GOkv+QjJD/jIyM/4qJi/+HhI3/goGZ/2xppP9EP5f/Ojml/zk5qv84OKr/OTmq/zo6rf86Oq3/Ojqt/zs6sP88ObLtNjSidjEvogBSUZ8BVFOhAGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAj47bAI+O2wJ3V/8Ah3r9K31s+NFuXvP/bFzz/2tb8v9sXPP/bFzz/2tc8v9pXfD/ZVjr/3Bm3/+8rPX//PD///rr+//+7/T//u/0//7v9P/87vP/+O3w//bq7v/05+z/9Obs//Di7P/u4ez/7N3r/+fZ6P/j0+j/28ju/8a3+/+ei+v/fW3s/25k3f9xbrD/gICo/4iFnf+Oi5f/kY6S/5GOkv+RjpL/kY6S/5KPk/+Sj5P/ko+T/5KPk/+RjpH/jIyM/4qKiv+Mioz/hYOZ/3JvrP9HQ5r/QD+g/zs7rP86Oq3/Ojqt/zo6rf86Oq3/Ojqu/j07sf9CPrjSOzesNS8lzwBFRZICRUWSAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAjo3bAIOB1ABJRbgA//9qAIlz/ZNwYvL2bV/0/2ha8f9sXPP/a13y/2pb8P9rW/L/Y1jo/3Jp4P/GuvX//PP+//rs+//+7/T//u/0//7v9P/+7/T/++7z//fs7//26e3/9ejt//Hk7P/u4ez/7N3q/+fZ6P/j0+j/28ju/8e4+/+lkvL/gG/x/3Rp3f9xbrD/gICo/4iFnf+Oi5f/ko+T/5KPk/+Sj5P/k5CT/5SSlf+UkpX/lJKV/5SSlf+TkJT/jo6O/4yNi/+QjpD/ioeb/3dzpv9NTJr/QUCh/zs7rP86Oq3/Ojqt/zo6rf86Oq3/Ozqv/kA+s/VPScCSBigAA09MoAFycL4Abmy6AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF9bvgBfW74BuJj/AIZv/Dh+bfnZcmb4/2lc8f9qXfH/aV/v/2xc8/9rXPP/Z1vq/3lv3//Rxvn/+vX+//zv+f/+7/T//u/0//7v9P/+7/T/++7z//fs7//26e3/9Obs//Pl6//v4uz/7N3r/+fZ6P/j0+j/28jv/8a4+/+kkvL/gG/x/3Ro3f9xbrD/gICo/4iFnf+PjJf/lJGU/5SRlP+UkZT/lJKV/5WUlv+VlJb/lZSW/5WUlv+Vk5X/kZGS/4+OkP+QjpH/iYab/4B9oP5XUZj/QkGh/zs7rP86Oq3/Ojqt/zo6rf86Oq7+PTyx/UA/tM06Na0uLim6AEJCjwJKSZYAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF9bvgBfW74AYVzBAYlx/wCIdP19dWf39mtg8v9pX+//al7x/2xc8/9rXPP/aF3q/4F33v/Y0P7/+vX8//zv8//+7/T//u/0//7v9P/+7/T/++7z//fs7//26e3/9Obs//Pl6//v4uz/7N3r/+fZ6P/i0+j/2srr/8m8+/+kkvL/f2/x/3No3f9xbrD/gICo/4iFnf+PjJf/lJGU/5SRlP+UkZT/lJKV/5WUlv+VlJb/lpSX/5WUlv+VlJb/kZKS/5CPkf+QjZD/ioeU/4GBo/9lXZX/RUKg/zs7rP46Oq7+Ojqt/zo6rv4+PbH/RkS28UVBt2k3MbEAXVnAAUdHmABCQo8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74AXlq+AYWD1gGIdvoWe2z4onRo+fxrXvH/alvy/2pb8v9sXvP/aV/r/4R64//e1///+vX9//3v9f//8PT//vD0//7w9f/+7/X//O7z//nr8P/36e7/9Obs//Lk6//v4uz/7N3q/+na6f/k1ej/3tDn/83D+v+qmvP/hnHy/nps4v93cbL/hIOk/4+Inf+PjJT/lJKV/5SSlf+Vk5b/lZSW/5WVl/+WlZf+l5WY/paUl/+VlJb/kZGS/5CPkf+QjpH/i4mM/4WEkv9qZpH/RkSg/zw7sP49PLH+Ozqu/j49sf9HRbb6UE68kmhiwgxdWb4B"
    s = s & "X1u+AV9bvgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74AambFAI+O2wGVt6cBhnj/I39w/rFvYfb/aVrx/2da8P9sXvP/al3t/42B6P/m4f//+/X9//7w+P//8Pb///D1//7w9//+7/b//ezy//zr8v/36O7/8+br/+/i7P/u4ez/7N3q/+vc6v/n2ej/49Po/9PF9P+xpPL/hnLv/ntt4v94cbH/ioWe/5iMnP+RjZD/mZaZ/5qWmf+alpr+mZeZ/5mYmP+Yl5j+lpaX/paVl/+VlJb/kZGS/5COkf+Sj5P/jIuN/4uKiv90cZH/RUKb/zw6sP4+PbT/PTuz/0hGu/BPTr2SVGLPFi0yaQFgXL8BX1u+AF9bvgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjo3bAI6N2wCOjdoBAAAAAIVu+zB2afmjcWH292pb8v9rW/L/alzu/5OF7P/p4v//+/L9//7w9//+7/b//vD3//7w+P/97vX/++r1//vr9f/35+3/8ePr/+/h7P/u4ez/7N3r/+na6f/m2Oj/39Ho/9jK9P+xpPL/f3Hs/3Zp4f93cbL/jIep/5iMnf+Xjpn/oJqh/6CZof+gmaH/n5qh/52bnv+bmZr/mpeZ/5eWmP+VlZb/k5KT/5"
    s = s & "KQkv+UkpX/jY6M/4yOjP98epX/SUaW/zw9sP49Pbf/QD234VNNw3lVYc8VJRYpAUJCjwJXVLIAX1u+AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wCWluAAWFe0Abl4/wCGeP0gemz5eHhl981uYPX/alzv/5aH8P/x5P7//fD6//7w9//+8Pf//vD4//7w+P/+7/b//Ov2//zr9f/36O7/8eTr/+/i7P/t3+v/69zq/+fZ6f/m2Oj/3NDi/9zQ7/+4qvL/kH3p/4t54/+Igrb/nZe0/6GYrP+imqv/o6Cl/6Sgpf+koaX/pKGl/6Ogo/+gnKD/n5qf/5uamv+YmJj/l5WX/5WTlv+VlJb/j46P/5KRkv+CgJv/VFKe/0lHu/RHR8CvQ0C3U1pWuAz//9MAQ0OPAkJCjwBCQo8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74AX1u+AF9dtwEsAP8Al5T5CIly/D91ZviTb2Ht36SS7f/67P3//fD1//7x9v/+8ff//vH3//7x9//98Pb/++3x//jq7//15u3/8+Xr//Di7P/s3+f/693p/+bZ4//l2OH/3dHd/9jM4f/Kvub/vKvn/7Gm3P+u"
    s = s & "qcn/uLK7/7Wwuv+zrrX/r6yu/6yqqv+qp6j/qaan/6akpP+koaP/oJ6g/5ybm/+bmpr/mJiY/5WVlv+VlJb/kZCR/4+Pj/+FhJv+WFapwlVPwWZcVs4oLSqiAAAAAACPjtwBMzSAAEJCjwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX1u+AP///wCNjNgBya//AKB//wCAaPgVgXHrS8u58aj67Pz//PD0//7y9v/+8vb//vL2//7y9v/98PX/+e3x//bq7v/16O3/8+bs//Lj6f/u4eP/693o/+XZ4P/k2N3/3dHc/9XK1P/VytP/0MfU/8jBzP/Fv8n/wry//725u/+6trj/tbK0/7Kur/+vqqz/qqeo/6akpf+loqP/n52e/5uamv+bmpr/mJiY/5WVlv+VlJb/kJGR/4uMi/+LiJj9YWO2aFRizgZ7bf8AAAcAAF5avgGhoecAjo3bAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wCfn+0AYmC0AJ2a0AH///AB///0CO7j+XX68v3//fH1//7y9v/+8vb//vL2//3x9f/87/T/+O"
    s = s & "zx//bp7f/05+z/8+Xr//Di5v/u4OL/69zh/+XZ3v/i1tr/29DW/9fMzv/Wy83/z8fJ/8vDxv/FwMH/wby//724u/+6tbf/tLG0/7Kur/+wqq3/qqeo/6akpf+loqP/n52d/5uamv+bmpr/mJiY/5WVlv+VlJb/k5GU/4+Pj/+Oi5H9gH2cbnV0oglvbr8DcHC9AF9bvgBfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGhmtABoZrQAf3zFALq16gCenOQC5drpEPHm8If+8/f//vL2//7y9v/+8vb//vL2//3x9f/77/T/+Ovw//Xo7f/05uz/8+Xr/+/i5f/t4OL/59ve/+TY3f/d0tb/2s/T/9fLz//Wy87/z8fK/8W/wv/FwML/wby//725u/+6tbj/tLG0/7Kur/+wqq3/qqeo/6akpf+loqP/n52d/5uamv+bmpr/mJiY/5aWl/+VlZb/lZOW/5CRkf+PjY/9kZCVh5KSmQyQj+gCkI/dAJCP3QAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAA1s7xANbO8QHLw/AD+e/1FPvx9pj/8/b///H1///x9f//8fX///H1//7w9f/+7/T/+evw//Xo7f/16O3/9OTp/+/h5P/o3N//6Nrd/+PW2f/YzdH/2s/S/9PKzf/Wy87/zMTI/764u/+7tbn/vbi6/7u2uP+zrbD/r6us/6+rrP+wqq3/rKmq/6mmp/+loqP/oJ6e/52cnP+cm5v/mJiY/5aXl/+Wl5f/lZWW/5ORlP+QjpD+mJiZmZudnQ+MiucCjo3bAI6N2wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA49r1AOPa9QHb0/QE/fP6Ff70+Zv/8/b///H2//7w9///8fX///H1///w9f/+7/T/+evw//Xo7f/05+z/8OLm/+zd3//m2d3/3NDU/9zQ2f/ZztP/vLW4/5SOlv+EfIL+e3B7/3Jndf93bHf/eXB5/4GAgv6Sj5P+oqCg/6ajpP+opaf/p6Wm/6elpv+mpKX/pKGi/6Cenv+enZ3/m5qa/5qZmf+XmJj/lpeX/5SSlf+QkJH+l5mXsZqcmhJiXuIBeHfFAHh3xQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAp6PjAKej4wBsbtIC/vP8F//z+qX/8/b//vD1///x9f//8fX///H1///w9f/97vP/+Ojv//Tn7P/x5Of/7eDi/+jb3v/f09z/1Mra/7Gqtv+Bd4H/WVFZ/z89SP8zMTj/Mi83/zc0QP8yLjb/Mi44/zk2SP9RSlL/X1xk/3R0df+PjY7/oJ6e/6Shov+joaL/o6Ch/6Ogof+gnp7/np2d/5ybm/+ampr/l5iY/5WUlv+TkZT+k5OT05KTmRmytUYAS0iWAF9dqwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAzMjlAMzI5QD///8A//T3G//097f/8/b//u/0///x9f//8fX///H1///w9f/97fP/9+ju//Tn7P/w4+b/693f/+TX2f/c0Nr/lZCp/1ZObv82MUv/PzZX/0pKc/9ZVYH/Yl2H/2Neif9aVoL/V1J+/0g+Z/88NFT/Ly1C/zw6Q/9ZVFn/gICB/piXl/+fnp7/np2d/6GfoP+joKH/oJ6e/56dnf+cm5v/nJib/5eXmP+VlJb/kpKT246OnyBZV/8AUk6kAGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/fbuAP//7wD/9fgA/vT2HP/z9rv/8fX//u/0//7w9P//8fX///H1//7v9P/86/L/9+ju//Pl6//v4ub/7eDi/93R1P+RiY/+QT1T/z41Wf9YS4b/dmm2/3tw1P+Fdd3/h3bf/4B02v+Fdd7/f3Pa/39v0/9eVqv/ODR5/y0rP/8xMT3/R0hP/3x8f/+bmZv/nZyc/6Gfn/+joKH/o6Ch/5+env+enZ3/nJqb/5eYmP+VlZb/ko+T9ZGRnCyPmsABZmS1AWdltQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7ujrAO/p6wD88voA/PL7IP3y+dD/8fX//u/0//7w9P//8fX///H1//7u9P/86/L/9+ju//Pl6//w4+b/59vf/5eSlv5FREz/Rj5i/3Blpf+Dds/+f3He/3No2v9zaNr/gHHp/3ls4/+Ccuv/gXHq/35v6P+Bdd3/em/H/ktOif8eKk//MC9G/1xaav+OjJD/n52e/6Ogof+joKH/o6Ch/6Cenv+enZ3/nJub/5eYmP+Wl5f/ko+T+YmCuz1+cu8DXlu9Al9bvQAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7ujrAOPb9AD78fMA+fDzJPvw9N3/8fX//u/0//7w9P//8fX///H1//7u9P/86/L/9+ju//Pl6//26e3/tayy/05MVv9EOVL/c2Wr/4975/+GeOP/kYPi/5SG4v+Qg93+j4He/pOF4/+UhuP/kIPd/pGE3v6MfuP/eGvj/3dr1v9MT5n/ISxN/z0+Rv9vbnL/mpWc/6Ofo/+joKH/o6Ch/6Cenv+enZ3/nJub/5eYmP+XlZj+kJGR+pmZnlKenqkFYl/DAmVhwgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7ujrANnN/QD///MA9+70KPzx9N7/8fX//u/0///x9f//8fX///H1//3u9P/86/L/9+ju//Tm7P/b0dX/dmp7/0I2WP9jXJn/fXPW/3ls4/+LfeD/0sr5/+jh/v/s5P3/z8X8/93W/v/o4f3/5t79/+Td/f+5rvf/fG3t/35u8v98btT/S0WB/yooO/9VU13/kouT/6KdoP+joKH/"
    s = s & "o6Ch/6Cen/+enZ3/nJub/5eYmP+Xlpj+kJGR+4iFlXaFgJkJjYzdAo2M2gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7ujrAPTv5gD///8A+/L5Kv/0+OX+8fX//u/0///x9f//8fX//vD1//ru8//47PD/9ejt//bo7v+up6r/SkhQ/1BGef95bM//fnHk/31v4v+Pg9j/4Nn5//nw///39P//3dX8/+vm///59v///Pr///bt/f/CufX/emzn/3lq9P99b+f/a2Gw/zs1V/9BO0z/hX2H/6SgpP+joKH/o6Gi/6Ogof+gnp7/np2d/5ybm/+bmpr/kZKS+5WWl4SXmJsKjozfAo6N2wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7ujrAOvi9QD///8A+/H5Lv7z+O/+8fX//u/0///x9f//8fX///H1//3w9P/47PD/9+ru/+/i5f+OiIv+S0JW/2RZov98buT/fG3u/4By5f+QhNj/5d75//v7///7+v//6eH+/+"
    s = s & "zl/P/59v///Pv///Tw///CufX/e2zn/3hp9/96a/D/eG3P/05KfP8wKjv/b2ht/6Keof+lo6P/pqSl/6Sio/+joKH/oJ6e/56dnf+bmpr/lZOV+5aWmISXmJsKjo3fAo6N2wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA7ujrAPnw9ADz6+oB/fP4Nv7z+O/+8fX//u/0///x9f//8fX///H1///x9f/77/L/+u7y/9fP0v9uZ2v/SEVk/3puyf99bvX/dmj3/35u7v+Pgtv/5N35//v7///7+///7OX9/+ba6//59v7//Pz///Tw///DuvT/fXDm/3hp9/95a/D/hHXk/21gqv8yK0L/VVRV/5uZmv+mo6T/paKj/6akpf+koqP/o6Ch/5+env+cm5v/lpWX/JeYmYiYmZsLhoTaAoeG1wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA7ujrAOLg4gDz7OoB/fL4Nv7z+O/+8fX//u/0///x9f//8fX///H1///z9v/8"
    s = s & "8fX//fH2/8e+w/9hWl//VU9v/4Byzv95a/b/b2T4/3lq7/+MfeP/3dX7//n3///59///6OL9/93T7v/39P//+vj///Pu///DuvX/fG/m/3hp9/95avP/hXPv/31qyf9AOFz/TkxM/5qXlv+ppqf/paKj/6akpf+mpKX/pKGi/6Cen/+enZ3/l5iY/Jqbmp6am5sNWFTAAWBcvAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA7ujrAOzn6gHy6uoB/fP4Nv7z+PD+8fX//u/0///x9f//8fX///H1///y9v/+8vb//fH2/7+4vP9cVlv/W1Rz/4N1z/95a/X/b2P4/31s9v+MfeL/z8b6/+nh/v/q4fn/39j+/97W/f/s4vj/6+L5/+be/f+8sfP/e2zp/3hp9/95avb/iXX0/4BszP9AOFz/TkxM/5yZmf+rqKn/paOk/6akpf+mpKX/pKGi/6Ogof+fnZ7/mJmY/ZmbmLCZm5gOUk7AAGhntAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAADu6OsA7ujrAOzn6gHv6OkB/vP4N//0+PP+8fX//vD0///x9f//8fX///H1//7x9f//8fb///P4/8S8wf9fWV7/WlRz/4N1z/95a/X/d2r9/4Bw+v+Ke+D+0MX5/+Ha/P/j2vD/5979/+LY8//p3u7/49rw/93W+/+2qfT/fW7v/3dp+v95avr/hXPw/3ppxv9BOF//TkxP/5yYmP+tqqv/pqSl/6akpf+mpKX/pKGi/6Ogof+fnZ7/nJub/ZeYl7uWl5kQkpSFAGtquAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA7ujrAOzn6gHv6OkB/vP4N//z+PP/8PX//vD0///x9f//8fX///H1//3w9P/98Pf///X//8vDyv9iW2H/WlRy/4N10P96bPX/eWz7/3ts7v+MfeL/7eH7//z5///69///+fH8/93S5f/8+///+vj///f1///JwPr/fGzt/3dp/P95avr/hHPv/3tqxP9COWj/UE5Y/5yYmP+qp6j/p6Sl/6elpv+mpKX/pKGi/6Ogof+fnp7/np2d/pOTk9aPj54XcW//AIeG0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA7ujrAOzn6gHv6OkB/vP4N//z+PP+8Pb//u/1///x9v//8fb///H2//3w9f/97/b//vb//9TQ2P9tZm7/WFNz/31xyv9+bvH/eGn2/3ts6v+NfuL/7N76//z7///7+f//+Pf9/+PX6//49f7/+/n///b0///FvPX/em3j/3pq9/97a/X/gnPk/3Jkrv86Mlr/V1Vh/6Cdnv+qp6f/q6eo/6qnqP+ppqf/paOk/6Ogof+fnZ7/np2d/pOTk9SNjKQbe3npAIOCygAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA7ujrAOzn6gHv6OkB/vP4N//z+fP+8Pf//u/1//7w+P/+8Pj//fD3//zv9v/98Pj//PP7//Lo7f+UiZH7VUxo/HNos/9+b+n/e2v0/31u5v+LfN7/5Nn6//r4///7+f//9/D9/93S6f/w6v7//Pv///Tw///EvPP/e2/c/35u8P+Ccez/iXrZ/19Vjf8wKkP/a2dx/6ilpv+uqqv/ramq/62pqv+tqar/qqeo/6WjpP+fnp7/np2d/pWUluOOjaUcdXP0AGlntQBoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA7ujrAOzn6gHv6OkB/vP4N//z+fP+8Pf//u/1//7w+P/+8Pj//fD3//3w9//98/v/+/D4///1+v/Bub36YFFk/GRUkP+Bctr/f3Dr/31u5P+KfN7/4tr6//n2///7+f//+fH+/+PZ9P/u6P//+/r///bz///IwPb/fG7g/39u7/+Hdur/hHXF/0hDaf9DP03/jYeO/6+rrf+tqar/ramq/62pqv+tqar/qqeo/6ekpf+hn6D/np2d/5yanO+Xl6Qih467AWhmtABoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA7ujrAOzn6gHx6uoB/vP4Nv7z+fH+8Pf//u/1//7w+P/98Pf//O/2//3w9//98/v/+/H4//3y9//o3uH/hXmE/1JGcf90arT/hnnd/4Bz4v+Lfd7/29P5/+vm/v/w7f//8ev//97V///g2f//7+r//+vl//++s/j/d2nl/4Bv7/+EduD/XlaU/zY0R/9qaW7/rKas/7KtsP+sqan/ramq/62pqv+tqar/qqeo/6qnqP+lo6T/oZ+g/6Gen+ibnaImjpioAWhmtAFoZrQAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA7ujrAOzn6gHz6+oB/fP4Nv7y+e/+8Pf//u/1//7w+P/+8Pf//fD2//3w9//+8fn//fH4//3u9v//7/T/yb3B/2Rdbv9TTHn/dm65/4d73f+FeN//n5Tr/6SY7f+nm+7/qp7u/6KX7f+il+3/qaDu/6ec7v+Thez/fW7q/4Jz4f9mX6P/NzZQ/0ZHUP+VlJX/s6y2/7Cqs/+xra7/r6ut/66pq/+tqar/ramq/62pqv+qp6j/pKGi/6Gen+mbnaIljZepAWhmtAFoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA7ujrAOnk5wD07esB/fP4Nv7y+e/+8Pf//u/1//7w+P/98Pf//fD3//7w+P/+8Pj//vD4//7w+f/66/H/9ebr/6+nrf9aU2X/VEx6/3xyvP+Hetz/fG7j/3xt5/98bOj/eWrn/3tt5/9+b+f/emzm/31v4/+Gdef/hHXe/2hdrf8+Olv/PDxD/4KBg/+yrbD/sauz/7Sttf"
    s = s & "+2r7T/s66x/7Crrf+vqqz/ramq/62pqv+tqar/pKGi/6Gen+mbnaIljZepAWhmtAFoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA7ujrAPDo8gD///8A/PL4Lv/0+ej+8Pj//u/1//7w+P/+8Pf//fD3//7w+P/+8Pj//vD4///x+f/56/D//O3x//Tp7f+po6j/WFFj/1dKcP9iXJn/enDD/4d44P6HeOL+gHLm/35v5/+CdOb/hXbf/oFz0f92bLv/WE+F/z87Wf9QS1r/g4GE/7Csrf+4srX/t7G2/7awtf+1sbX/tbC0/7Ousf+xrK7/sKqt/62pqv+tqar/qqan/6KfoOmbnaImjpioAWhmtAFoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7ujrAN/X8wD/+PsA/PL7I/3z+tf/8Pb//u/1//7w+P/+8Pj//vD4//7w+P/+8Pj//vD4///y+P/+8f"
    Sec08_HeaderLogo_7 = s
End Function

'--- Sec08_HeaderLogo_8 : Dashboard header, toolbar buttons (Layout/Update/KPI/Chart/Slicer/Save/Reset/Deploy) and the embedded company logo bitmap (chunk 8 of 18)
Private Function Sec08_HeaderLogo_8() As String
    Dim s As String
    s = s & "X/++7x//ns8P/s4eb/ubC0/3Rsc/9OSl3/UEdy/19bl/9sX6r/bme4/2xhvf9vZ7b/a12l/1pRhP9JR2b/RD5M/2FhZP+Wk5f/s7Cy/7m1uP+4tLf/vLi7/7q1uP+2s7b/tbG1/7WwtP+yrrD/sq6v/7Crrf+wqq3/rqmq/qmnpd6enqwffoPBAWhmtABoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA8OrpAPDq6QD/+P8A/vT4Hf7z98D+8fX//u/1//7w+P/+8Pj//vD4//7w+f//8fr///H6///z+P//8/b//e/z//rq8P/46vD/9+nt/9rMz/+knqH/dG6A/1lQaP9QQGb/UUNq/1BBbP9JQGv/UUFl/1dLYf9iXGf/ioiJ/6ilpv++urz/vrm8/8C6vf+/ur3/wbu+/766vP+7t7r/ura4/7eytv+0sbT/srCx/7Gsrv+vqqz/ramq/q2rqNmnpa8XAAD/AGlntQBoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA49r1"
    s = s & "AOPa9QHZ0PMD/vT4Fv7z+KL98fX//O/1//7w+P/+8Pj//vD4//7w+f//8fr///H6///z+P//8vb//vD0//3s8//36e7/9uju//fp7v/t4uX/3dLY/8O3vf+imqH/j4iP/42Gjf+IgYz/mJGY/66lq//Btrr/zsbI/87FyP/HwcT/w7/B/8O+wP/CvL//wry//8K8v/+/ur3/vbm8/7u2uP+2s7b/tLG0/7KvsP+uqqv/ramq/aGhnrqen54Qn6CXAI6N2wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAq6jkAKuo5ACNjOAC7ePrEvXr8I//8/b//vD2//7w+P/+8Pj//vD4//7w+f//8fr///H6///x+v//8ff//vD0//3s8//36e7/9Oft//Tk6f/05Oj/9efs//Hi5f/q297/5NbZ/93R0//f09X/39LU/9vR1P/b0NT/2s/S/9TKzf/Oxsn/zMTI/8rCxv/HwcT/w73A/8K8v//CvL//vrq9/725vP+7trj/urW3/7awtf+yrrD/sa2u/J6fnp2ZmpoNjYzhAo6N2wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAxsPkAMbD5ACwr94B6uLrBfru81f/8vX//vH2//7w+P/+8Pj//vD4//7w+f//8fr///H6///x+v//8Pf//e3z//zr8v/36e7/+erw//vq8P/05ev/7+Hl//Lk6f/v4eT/7N3f/+rb3f/j19r/4tbZ/9/T1v/ZztL/18zP/9bLzv/Tycz/z8fK/87Gyv/Oxsn/ycLF/8S/wf/FwML/wr3A/8G8v/++ubz/vLe5/7aytv+1sLT/tLG0+56en4eWl5kKjo3fAo6N2wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//ruAP/+7wD//9wA////AP3y+DT+8/fp/PD1//7w+P/+8Pn///H6///x+v//8fr//vL7//7y+//97vX//Ovy//zr8v/76/L//Ovy//fp7//05uv/8uLm//Pi5//y4uX/7d/h/+rb3f/o2tz/49fb/+DU1//f09b/3dLV/9jN0f/Wy87/08nM/8/Hyv/Px8r/zsfK/8zEx//Kw8b/x8HD/8W/wv/Cvb//wbu+/725vP+7trj/trO2+aenqFafoKEF0s3rAc7K5gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7ujrAJKR3ACMi9sAAAB5Af7y+x3+8vux/PH8//zv/P/+8fn//vL7///x+v//8fr///H6///x+v/+8Pj//u/1//7v9P/+7/T//ezz//vr8v/26O7/8+Xr//Pl6v/z4uf/8uLl/+3f4f/q293/6Nrd/+PX2//g1Nf/4NTX/97S1v/az9P/1szP/83Fyf/Nxcn/z8fK/83FyP/Lw8f/x8HE/8O+wP/Au77/vrq9/8G7vv/Bu77/uLfD27e2xCK9tsYA///1AP327gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wCOjdsAUFnNAdnL6hDo2vWE9un+//zv/f/+8Pn//vH7//7y+///8fr//vL7//7y+//+8Pn//vD4//7w9v//8fb//u/1//zs8//76/L/9uju//Xo7f/z5ez/8+Ln//Li5v/s3+L/6tve/+ja3v/k2N3/5Nfd/+LW3P/d0tb/2s/U/9fN0f/Wy8//08nO/8/Hy//Nxcr/x8HF/8S/w//Oxcn/yMDG/7u2vP+wrbP/eXya31xhjiMAAV8AVluQAO7o6wAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIuK2QCLitkAj5L/AXx4jhKdk8OLybr2/+7e/f/y4/3/8+T9//Tl/f/26f7/+Oz///nt///57v//+u7///ru/v/67v7/+u7+//jq/v/46f3/9+n9//Xn8f/x4vv/7uHu/+3g7f/q3Ov/6dvq/+XY6f/k1un/4dLp/97Q6f/YzOP/2Mzh/9fK4f/Vyd3/08jg/83D2//IvN3/vbTX/7ax0f+oorz/h4at/3Ftnf9YWIn/QTyO4ExHnSHo7/8Arav/AI6N2wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGJfwABiXsAAUFSnAYZ19AeLee1fkX7p+5+K6/+pl/L/tKXy/8Cz9v/Ivfv/zcP5/9bN+//d1Pv/4Nb6/+DW+v/f1vr/5tj6/+jY+//o2Pr/6Nj6/+fY+v/f1fn/3dT5/9fP+f/Qx/T/z8b0/87G8v/OxvL/ysH0/7+08v+1p+v/saTq/6KV4P+SicT/hYOw/3Zzr/9rZqz/XVmo/1BOov9J"
    s = s & "R6b/RkWo/05Kt/5PS7r7Uk7CeldTxwuUlNwCjIvZAIyL2QAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF9bvgCgoOYAhoXVAYfrSwCJcvogi3H2rY1y8/+Oc/X8jnfx/o577P+Pfuj/loTm/5yG6f+ci+H/nY7c/52N3P+djdz/m43b/5uN2v+bjdr/m43a/5mL2/+Shdn/kYTY/4+D1v+Gd8z/g3bK/391vv97cb//cGfE/3BnxP9tY7//Z120/2Zcs/9mXLP/Xleu/1NPtf9RTb3+VlHG/1lTyv9UUNL+Uk7U/lJO1P5RT9TkVFHWL25n3wFubbcBa2m3AGtptwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABgXL8AX1u+AFVWqwF0ddsFiHHwWo5z9/eUeP3/lnv//5d8//+Xffr/ln71/5Z+9f+XfvX/l3/1/5d/9f+Wfvb/jnru/4t56/+Meev/i3ns/4Rz7/+Ccu7/f3Dn/4Bx5/+Id+z/gX"
    s = s & "Hv/39w7P98buj/emzt/3pr7P96a+3/dmnp/3Jm5P9yZuP/bGLi/2Nb3v9hWdv/Ylrc/2Ja3P9bVeD/VFDa/lJO1v1MSsGNSUi1EKSj5QKKidcAmpnnAGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74Ai4rYAIuK2AGFAP8AiXL6IIty/aiPdP78lHf9/ZV6/v+Wev//lHj+/5R4/P+Wev7/l3z//5Z8//+Pef7/jXj+/414/v+MeP7/jHj+/4l2//+Jdv//hHP//4Rz//+Ecv//g3D//3xu/f98bv3/em38/3hq9/92aPf/b2P3/2tg8v9nXO3/YVbo/1xU5P9ZU+L/XFTk/1pT4v9XUt//V1Lf/1JQ1uNTUNUyg3nrAW5tswFnZbMAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAZ2WzAGRhrwBcV4QAhLHoAoZ2/EmMdPzbk3f8/5N2/P2UeP3/lnr//5d8"
    s = s & "//+XfP//l3z//5B5/v+NeP7/jXj+/414/v+Kdv//iXb//4l2//+Jdv//hHP//4Rz//+EcP//hHD//4Rw//+CcP//fHD//3ls+/92aPf/bmL0/2lf7/9pXe//Zlns/2BW5/9eVOb/WlPj/1hT4f9UT9j+Uk7V/k1LwotIR7IRvLvwAYuK2ACHhtQAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AIqJ1wCPjtwBUwD/AINs/BSMc/yNkXb9/pJ2/P2Td/z+lHj+/5Z7//+XfP//l3z//5J5//+NeP7/jXj+/414/v+Kdv//iXb//4l2//+Jdv//hHP//4Rz//+EcP//hHD//4Rw//+CcP//e2/+/3ps+/92aPf/bmP1/2lf7/9pXe//Zlnr/2BW5/9dVOX/WVPi/1hT4f9VUNv/U1DW2lVS2TJbT6sBXVygAWhmtABlY7EAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAI6N2wB7ec8AX1TJAIx0/wCHc/02inP+0JB0/v+Ud/39lXn//5Z6//+XfP//l3z//5Z7//+Ref//j3j+/494/f+Ld/7/inf//4l2//+Jdv//hHP//4Rz//+Ecf//hHD//4Rw//+Cb///em39/3ps+/92aPf/bmL1/2lf7/9pXe//Zlns/2BW5/9aU+P/WFPh/1hT4f9ZVOP7WVTXcVZTwQszNmcBQECOAEhIlQBoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74AXFi8ADJdVgCHcPkIi3L+bY50/vOSdvz/lHj9/pV5//+We///l3z//5d8//+Wev//lHj9/5R4/f+QeP3/j3j9/4t3/v+Kd///hnT//4Rz//+Ec///hHH//4Jw//+AcP//em39/3hq+/90Z/f/bmP1/2le7/9oXO7/ZVnq/2BW5/9aU+P/WFPh/1hT4f9bVuS9XljoIkQ85QC1sv8AXV2qAEJCjwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74AfHrOAIqH2AGLYP8AiXH3H4py/KKOdP79knb8/ZR4/v6We///l3z//5d8//+We///lXn//5R4/v+Td/z/k3f8/454/v+NeP7/iXb//4Rz//+Ec///g3P//31x//98cP//eWz9/3Vo+/9xZfj/bmP1/2hd7v9nWu3/ZFnq/2BW5/9aU+P/WFPh/1ZU4OtUWNhSTmepBJ+X4AGPjtwAjo3bAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYF6sAFRRnwB+eNIBi2D/AId1+TuLc/7MkHX9/5R4/v6We///l3z//5d8//+We///lXn//5N3/f+Td/z/k3f8/454/v+NeP7/iXb//4Rz//+Ec///g3P//31x//98cP//eWz9/3Vo+/9wZPj/bWL0/2hd7v9nWu3/ZFnq/2BW5/9aU+P/WVTj+lVX4YxMW9wUO1j/AFVStQB3dcsAjo3bAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AIaF0wCPjtwBk57HAYiD8QSKcv5kkHX96JR3/f+Vef/+l3z//5d8//+Wev//lXn//5N3/f+Td/z/k3f8/454/v+NeP7/iXb//4Z0//+Ec///g3P//31x//98cP//eWz9/3Vo/P9vY/f/a2Dz/2ld7v9nWu3/ZFnq/2BW5/9aU+P/W1Xmwl9W6jEAAP8AgILEAV9dpwBpaLMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wCOjdsAj4zfAYZf/wCAa/oRjXP8g5J2/O+Ud/3/lXn9/pZ6//+Vef//lHj+/5N3/P+Td/z/k3f8/454/v+NeP7/iXb//4l2//+Ec///g3P//31x//97b/7/eWz9/3Vo/P9wZPf/a2Dz/2hd7v9oW+7/Zlrs/2BW5/9YVeLoVlbhYGNZ6gfa4bAAkpHgAIWD0QBoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACOjdsAfXzKAGtkwAGMc/8Ah3H1HIx0/KaRdvz9k3b8/ZN3/f6Vef//k3f9/5N3/P+Td/z/k3f8/414/v+NeP7/iXb//4h2//+Ec///g3P//3xw//96bf3/eWz9/3Vo/P9wZPf/a2Dz/2ld7v9nWu3/Ylfp/19V6PtZV+SHRFzSEElW+ABbV7sAd3XKAI6N2wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AFRSoACIhNoBiF7/AIJy9TiQdvvLk3f8/5N2/P2Ud/3+k3f8/5N3/P+Td/z/k3f8/5F3/f+NeP7/iXb//4h2//+Ec///g3P//3xw//96bP3/d2n8/3Jm+f9wZPf/a1/y/2hc7v9mWuz/YVfq/19W56xdVd0jRUXxAI"
    s = s & "CA1AFSUJgAamiyAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AISD0QCPjtwBlZXeAUR3ewKOdPtOknb82ZN3/P+Td/z+k3f8/5N3/P+Td/z/k3f8/5N3/P+NeP7/iXb//4h2//+Ec///g3H//31w//95bP3/d2n9/3Rn+v9wZPf/a1/y/2da7P9iV+n/YFbpyWBY5j4AAAAAZWW0AqGg7gCEg9EAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wCOjdsAjYvcAKS9twF+bfsEjHL8apJ2/OeTd/z/k3f8/5N3/P+Tdvz9k3f8/5N3/P+NeP7/iXb//4h1//+D"
    s = s & "cf//fW7+/3pu/v93afz/c2b6/3Jl+f9tYfT/al/w/2Zb7P9dV+XmVVfhZHFk8QV/hX4BMDB+AGNisACOjdsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACOjdsAjo3bAJKP4QDXf/8Afmr6DYxz+4WSdvzwk3b8/5N3/P6Td/z+k3f8/494/f+NeP7/iXb//4l1//+DcP//em39/3ls/f92aPz/cGT3/25i9v9rX/P/Zlrt/2Zb6+9kXOl+O1rSDDM/8wB/fdEBZGOxAEJCjwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjo"
    s = s & "3bAH580QBjW8cAjHL/AItz+hSNc/uKknb875N2/P6Td/z/k3f8/5F3/f+Ndf/+iXb//4Vz//+DcP//em39/3hq/f92aPz/cGT3/2xh8/9pXfD/ZFju+GFY6ZV5a+gSYWHwAF1bvwFbW40AXVupAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF9bvgBkYbgAa2a7AYZp/wCJcPsSjHP7ipJ2/PCTdvz9k3f8/5N3/P+Nc//+jHT//opz//6DcP//emz9/3dp/f91Z/v/bWD1/2tc8v9lWe73X1bqnVxW5BsRAP8AhYLYAT43qwBiX7sAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAbm6uAGBbvwGbev8AiXD7EYxy+4uSdvzwk3f8/5N2/PyOc//+jHP//oZx//+DcP//emz9/3dp/f91Z/v/bWH1/2pd8v5kWey0b2XmJEoj/wCQj9wCkZDdAHx60ABfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAY1+6AF5avwBoZbcBknT/AIlw+xCKcfxwknb86o50/vyNc//+jXP//oty//6DcP//emz9/3dp/f91Zvv+aVvx9GJZ66dzaOkotq3XApKS2gGMi9sAjo3bAI6N2wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX1u+AGFeuwBoZ7QAaGa1AQtIAACFb/sLjHL9SY5z/rGNc//wjHP/9oRw//l7bf35eGr9+Xdp/PR4aPnLfmn5X19a5BCVjN4CkJHZAo2L3ACOjdsAjo3bAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGhmtABoZrQAaGa0AGlntQCPkNcCiXL6BY50/hGJfPwxhnX9YIVv/XyEb/x9g278doNu+0V8b+wZi3L/Bk+kAABhXb4BpKXqAI6N2wCOjdsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAA//////+AAP////////////+AAP////////////8AAH////////////8AAH////////////4AAD////////////4AAD////////////4AAD////////////wAAB////////////wAAA///////////+AAAAP//////////8AAAAH//////////4AAAAD//////////4AAAAD/A//////AfwAAAAA+Af////+APAAAAAAcAP////8AGAAAAAAIAH////8AAAAAAAAAAH////4AAAAAAAAAAD////4AAAAAAAAAAD////4AAAAAAAAAAD////4AAAAAAAAAAB////wAAAAAAAAAAB////wAAAAAAAAAAB////wAAAAAAAAAAB////wAAAAAAAAAAB////wAAAAAAAAAAB////wAAAAAAAAAAB////wAAAAAAAAAAB////wAAAAAAAAAAB////4AAAAAAAAAAB////4AAAAAAAAAAD////4AAAAAAAAAAD////8AAAAAAAAAAD////8AAAAAAAAAAH////+AAAAAAAAAAP////+AAAAAAAAAAP/////AAAAAAAAAAf/////gAAAAAAAAA//////wAAAAAAAAB//////4AAAAAAAAD//////8AAAAAAAAH//////8AAAAAAAAP///////AAAAAAAAP///////AAAAAAAAP///////AAAAAAAAP///////AAAAAAAAP///////AAAAAAAAP///////AAAAAAAAP///////AAAAAAAAP///////AAAAAAAAP///////AAAAAAAAP///////AAAAAAAAP//////+AAAAAAAAP//////+AAAAAAAAP//////+AAAAAAAAP//////+AAAAAAAAP//////+AAAAAAAAP//////+AAAAAA"
    s = s & "AAP//////+AAAAAAAAH//////+AAAAAAAAH//////+AAAAAAAAH//////+AAAAAAAAH//////+AAAAAAAAH//////+AAAAAAAAH///////AAAAAAAAH///////AAAAAAAAH///////AAAAAAAAP///////AAAAAAAAP///////AAAAAAAAP///////AAAAAAAAP///////AAAAAAAAP///////gAAAAAAAP///////gAAAAAAAP///////gAAAAAAAP///////gAAAAAAAP///////wAAAAAAAP///////wAAAAAAAf///////4AAAAAAAf///////4AAAAAAA////////8AAAAAAA////////+AAAAAAB////////+AAAAAAD/////////AAAAAAD/////////AAAAAAH/////////gAAAAAH/////////wAAAAAP/////////4AAAAAf/////////4AAAAAf/////////8AAAAA//////////+AAAAB///////////AAAAD///////////gAAAD///////////wAAAH///////////wAAAP///////////4AAAf///////////8AAA//////KAAAAIAAAAAAAQAAAQAgAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    Sec08_HeaderLogo_8 = s
End Function

'--- Sec08_HeaderLogo_9 : Dashboard header, toolbar buttons (Layout/Update/KPI/Chart/Slicer/Save/Reset/Deploy) and the embedded company logo bitmap (chunk 9 of 18)
Private Function Sec08_HeaderLogo_9() As String
    Dim s As String
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX1u+AG1srwBqY7IAZWi3Af8AAABRYMkfUV3WXEpZ2J02UL+zMEmMiS9DY0IsN3gLLjt3AWFdwQJgXL8BX1u+AF9bvgBfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjo3bAI6N2wB6eM8AVlzBAf8AAAA/fMwCM47mLzZ914c4dNHMN2rK3i5dteUhTJDYGUF4uRY9cngfMmEVPyhVAqeT/wFeWr0BaGa0AGRhuABoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACOjdsAjo3bAJCP3AHuVqABMorvBDOP7R84ieOMNoDZ5DN40P8xb8X/KmGx/yBRl/8ZRoT/FkB52RY9cnQXOnQXADxDAGVivgF2c8IAd3XCAHZzwgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjo3bAI6N2wCOjdsAjI7cAgCm/wBAlOoUOZLqaTuI4eE3gNn/M3jQ/zFvxf8sY7X/Jlqk/iBQlP8ZRoT9Fj103iM9bFVDQF0OJzRgAFNSoAFRUZsAVVSmAF9bvgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACOjdsAjo3bAI6N2wBMg9IA/wAAAEeS4zg/lOfPPIri+TeA2f8zeM//MnDH/y9pvv8rYrP/"
    s = s & "Jlmk/h5Mjf4ZO3P2KT5nuTJAZEQdTAAAQ0OQA0lImgAeI1QAX1u+AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wCOjdsAjo3bAg+K5gAyiuIeO47ldkOS6O9DkOf/O4Tc/zN3z/8ycMf/L2m+/yxjtf8jVZ3/GUeF/xhDf/4bPHH0ITh3fzozqRhIR5sBYFy9AWBcwABfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74AeHbOALG49QCPjdsBMZ/3AECY8EE/lOy6Porj+jyG3f83f9f/M3jP/zJwx/8vab7/LGK0/yRXof4cS4z+GEaE/xc7dP8jOm62Nz9sSgQgOQBrY9ICXFm7AWFdvABoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFpVuwBaVrsAXFC2AQAAcwAnhekIQ5fvbEOU7Ow+iuP9OIHZ/zV81f8zeND/MnDH/y9pvv8sYrT/J1ql/iBQlP0YRoT/GDt0/iA7a+UnPWeEHC5gCEVImwKHeugAgH7DAHZzwQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaWbEAGlmxAFtYsECYmzKADKU8Sw/kuqaRJHo/0SQ6P49ieL/N4DZ/zN4z/8ycMf/L2m+/yxitP8nWqX+IFCT/RhGhP8YO3T9GDlv/xk5b7UlOWMlJzljBPm1/wBYV6UBWFelAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACLidkAi4nZAZOI1gJ9i9wBOZrzUUCU7LtEkej/RJHo/j2J4v83gNn/M3jP/zJwx/8vab7/LGK0/ydapf4gUJP9GEaE/xg7dP0YOHD/HDtvyy9EaFgxRWgKeUD9AUNDkAJDQ5AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AIKBzgCWluAAxn3MADA+6ABBmvJ3RZrx3Uma8f5JmfD9Q5Hp/jqD2/8zd8//MnDH/y9pvv8sYrT/J1ql/iBQk/0YRoT/GDt0/Rc5cP8fO23dM0FjgDRCYg8tQVEAQECNAVJRngBoZrQAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfXKoAX12rAAAAsgDRnE4BkpugCUqc8ZlJnfT4S571/Eud9PxGlO"
    s = s & "v9PYjg/zZ91f8zdMv/L2q//yxitP8nWqX+IFCU/RhGhP8YPHX9FzRp/xo2Zu4iP2ynJj5vGRNEYQAAAAAAdXPBAF5cqgBsargAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AGlosgBlY7cAX1u+AHd1xgB8esgBcGrMAaCStQKPlcIUT53yrkud9P9LnfT8S5zz/EaU6/4/i+T/OYPd/zR50f8wasD/LGK0/yVYo/4eTY7+GEJ//hc4bv4WLl3/FTFh/BY9dMkeOnEpYyBSAbuX/wBTUp8CY2KvAHh1wwBoZrQAaGa0AGhmtABoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AGhmtABoZrQAaWiyAGVjtwBfXL4A1Nj/AI+N3AGSkNMBlYT3AoaH9hZNm/OsSp3z/0ud9PxLnPT8RpTr/j+K4/84gdr/M3fP/y9pvv8rYbP/IlSb/xhDgP4YOXD9Fi5d/xUqV/8WMF//GDhs2hw1ZzMnKlEEdVnZAUZFkwNlYrEBb2y7AGdlswBoZrQAaGa0AGhmtABoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAaGa0AGhmtABpaLIAZWO3AV9cvgJeWr8A////AAAAAAAAAAAAk5L/BkKW6pZFmu/3S571/Eud9PxGlez+O4LZ/zJxyf8var//K2Kz/yZZpP4eTpH+GEOA/xg5cP0VLV3/FStY/xUrU/4WK0jLFyxNHBoqOQAjNTMAAEQAAP///wBtabkBZ2azAWhmtAFoZrQAaGa0AGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AGhmtABoZrQAZmSxAHBrrAAKEgQAJyoZAFFSOgAAAGADDB2IESgvnSEsQnE6VXiqplt8tvJXdLT+Tm2s/kxqqf4yWZr+KE6W/jtQpP84TKb/JEaL/RY8dP4ZNWT/JjRQ/icwSP8jKkD/HCc9+xcoPtEpMDhbMjM0SjE9UUcwSnYxHD6HDD5AQAAnOV0AeHa7AGhmtAFnZbMA"
    s = s & "aGa0AGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAaGa0AGhmtABpaK4BABMAACUmOwkrLTIsUE5USVRVXWZMUGF8Rkpmi0JMX6RcYHHVYmF49lhUcv9OSmr/Skdp/zdCbf43Q4D+UEqd/0lFoP8xPIr+IDBm/iAoRv8sKzT/Liwz/yooLf8kJS78Hyc26SkqMb0rKy+2Ky87sioxQ5wnKjh0LztaRzVHdxcbLAAAa2i0AXRxvwB4dcQAdnPCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF9bvgBfW74AYFy9AGZktgBgXL0AX1u+AF9bvgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAaGa0AGhmtABoZrQAaWezAGJeuwBfW74AX1u+AF9bvgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7ujrAHl3uwBXVa0AZmSzAQANHAAuP3gMNDxWOz0+Q5JUVFfNXFxf8lpbXv9WWVv/Vllb/1tdX/9aXF7+U1RY/1BLVv9EQlT/R0hp/1ZTjv9hWaz/VFCn/0VBnf80MG7/KSY5/yknL/8pJy7/KCYs/yckKv8mJCn/JiQp/yYkKf8mJCn/JiMo/yYlLPstNUvFMDpWWzAxRxD///8AWFalAVRToQBiYa4AY2GvAGhmtABoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEJCjwBCQo8ATEufAGJexABgXL0AZmS2AGBcvQBfW74AX1u+AF9bvgBfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGhmtABoZrQAaGa0AGhmtABpZ7MAYl67AF9bvgBfW74AX1u+AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA///1ANvW4gCXk8EAAhcgADNBZjs9R2CsVlhd7V9gYvtgYmT+ZGVn/2xtb/5s"
    s = s & "bG/+bW1v/mdoav5eYGL/WVld/k9LVf9MSmj/WFaS/2dfuv9RTbf+SEak/zg2bv8qJzz/Ly45/y8vOf8qKTH/JCIn/iAeIf0gHiH9IB4h/SAeIf0gHiH9IB4i/SYlLPorLDKuMDI2RgAEAABDQ5IDSEiVAQAAPgBjYa8AaGa0AGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQkKPAEJCjwAYHUoAXFi6AWBdvQJmZLYBYFy9Al9bvgFfW74AX1u+AF9bvgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGhmtABoZrQAaGa0AGhmtABoZrQBaGa0AGlnswFiXrsBXlm+AZ2d5AB9e9EAjo3bAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wDu6OsA7ujrAO7o6wDu6OsA7+nrAfv07wEAAAAAU1RkUlJUXdZaW178W11f/1tcX/9hYmX+bm5w/mZnaf9kZmj/YGJk/1lbXv5SVFj/RURT/0ZGZ/9YVpL/Zl+7/0hGtv9APaL/NTFt/yclOP8oJi7/KCcu/yMhJ/4fHCD9Hhse/R4bHv0eHB79Hhwf/R4cH/0fHSD9JSMo/ysrMLkwMjZOBgwAAEREkwRwb70DiYjXAWNhrgBaVcQAYl+6AF9bvgAAAAAAAAAAAAAAAAAAAAAAaGa0AGhmtABdW6kAAAAAAEFBjwJNS5QBIiWPACUkmgAAAI4AiYfPAVxYvQFoZ7QAY2"
    s = s & "C5AGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAaGa0AGhmtABpZ7UAYlzAAXVl+wBoXesAPlHJACVCrAAAAAAAj47cAZKR3wB7eccAaGa0AGhmtAAAAAAAAAAAAAAAAADu6OsA7ujrAO7o6wDu6OsA7ujrAefi6AD65f8A3c30AId/bABUUlhSV1VZ31dYW/teXGH/YWBk/mZmaP5xcHD/b29w/3JydP9ycnT/bW1v/2dnav9aW2//XVuG/2dhpP9mX7r/S0i2/0NAof89O3v/PTxY/0hES/9EQ0f/Pj5B/jk5PP43ODv+Nzg7/jYzOv4zLjX+MCow/i8qMP4rKTH/KisxtScrLkZyaG4AfXpsAAIdGAAAACEAdFb/AF9bvgFdWMAAYl+7AGhmtAAAAAAAAAAAAEJCjwBJSZYAm5fmAGhltAFqaLMBLCe6ACo0ewgxM5UpNTSeVjo2qTtAObMSJiaZAGhmtAAAAP8AYl+7AF9bvgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGloswBpaLIAaWizAGljvAF0Q/8AbWPkCmhd"
    s = s & "6jA9UsQ0KUSmHy0tewj///8AycruAFxZrQBzcL4ARkWTAEJCjwBCQo8A7ujrAO7o6wDv6ewA7ujrANfX0wDZz9MAvbK7AP///wLj0/sS////FIeEhmNwb3Hfc3F0+4F8gf+Hg4f+iYeI/o6Ki/+TkJH/nZmb/6Ccnv+cmZn/l5SX/4+Lov+Nhbr/hHnI/2xkw/9STrr+RkSk/1BQlP9paI3/f3iE/nh2e/9wcXT/aWpt/2RlaP9iZGb/X1xk/1lTWv9TTFL/UkxR/0VCS/9FQ0vMVVFXfnpueEFvbHgxQlFvIitAhxFPRf8D//8AACYo9QB2dKMAaGa0AExLmQBCQo8AQkKPAEBBjgFKSZgBlpSyADo7uQA8NK8OOjirQTg3qY04NqjJOzivpDs4rkUiL20FZmS1AWJevwFeWr8AX1u+AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAZmS2AGZktgBmZLYBpYj/AKiT5QV2Z+04aF7rr0VUzL44SbiAPDupNDw3rQpPSq8CgX+5AV5cqwFAQI0AQkKPAElJlgDh2+YA7ujsAO3n6gDu6OsB/fP7AtzU1wnf1tgW5driLuHT61zr3u1yzcbHo7KurOi1r7H9wry//8bAw//EvcD/w73A/8O9wP/Pxcn/08nM/83FyP/Jwcb/wbvR/7iu5v+fk+n/dWvQ/1NOvP5CQKj/YmCq/5SRtf+ln7P/opyn/5mWm/+Sj5L/jYmN/4eGh/+CgYL/fXp6/3l2df94dH"
    s = s & "T/cGpv/m1nbf1xaXH3d212429td7VNWG6JSFF7WGpinip/epgLfnqkAg0U/wBta7AASEiWADw9hQBCQo8AQkKQAkNDigJVVb0BSEe2DT46sE89OrO0PTq1/j06tf09O7X/OzquljU3mylqZMcAY1/BAV9bvgFfW74AX1u+AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGhmtABgXL0AYFy9AGBcvQKSdf8Aj3T8M4Bs95xqX+/uYFnn91dT3OhNSs6tPTiwTjs2rQs4M60ARESRAkJCjwNCQo8AQUKOAO/p6wDu6OsA7OfpAfXs8gH/9PwQ39baUNnQ1KTb0NTa39PY5uPX3Oni1tvx3dLW+93S1//k1+H/49ff/97T1//c0dT/2c7X/9nO1f/YzdH/08nM/8zDyf/DvdP/urDp/5yQ4/9qYsD/Uk25/kJAqf9iYar/lpK3/6ehtP+jnqn/m5id/paTlv6Tj5P/j4qO/4mHiP+DgoL/gH5//399fv95dnb/cnBy/29ucf5wb3L8b29y9GxtcexsaXDocWp2uoF8lFWEf5oNjIaKAGhmtAFMTJkAQkKPAEJCjwNBQY4CRkW+AEtKtg9JSLZpQkGzxz49tPg9O7b/Pj+4/j8/tP9APrTYQT20cUVCsgAAAIoAX1u+AF9bvgBfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYV+tAGFfrABiX64BX1u8AQAAAAF7afhceWj43Wxh8/5oXO//XVbl/1RQ3PBSTdW+SUi7ZDxAmBqGaf8A/96QAENDjwJTUqEBvLjWAf//AADz6fEC9uzzEvvx9G3v5Ou/6d3l7uzh5P/t4eX/7uLm/+3h5//s3+j/7+Hn/+7g6//t3+r/697n/+rd5v/p3Of/4NTe/9rP1//Yzdb/z8bV/8O91f++tOf/oZTo/2phzP9WUcT/REKs/15dpv+OjK7/o52r/5+VpP+WkJz/ko+X/5SRlP+QjZD/ioiJ/4SCg/+Bf4D/gX+A/4F+gf59e33/enh8/3p3f/96d37/d3V2/3Vycv90cHP4cnF3vnJzekNzdYAM//8AAEVFlQNeW7cBKSewAAAuAAMrMostODehgEJBps5AQbL3Pj+4/z4+uP0+QLn9QEC0/0FAtP0/PrGoLzOTHywxjgP/8/8AYl7DAWJewwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACBgM4AgYDOAYKBzAKJh9YAenH0FnJm9YFwYvX7bV/0/mlb8P9jWOr/Xlbm/1xV4vxYU87IUk7BhEdEsD4cJHANl4j/ABwkogDh2PYChYThAPzy+Ar78fdE+/D0vvbq8P305+7/+Ozw//fr8P/36u//9Oju//Lk7f"
    s = s & "/z5ez/8eTs/+/i7P/u4e3/7uHt/+7g7P/l2OX/39Lh/93Q4v/VyuH/ysHe/8S46v+mmev/cWfW/1pVy/9FQ7P/V1ak/358oP+SjqD/koqa/4yGlv+LiJP/j42Q/42Ljv+KiIr/hoWG/4WDhP+Fg4X/h4WJ/oaEh/+EgYb/hH+I/4WBiP+AfYD/fnt9/nx6fP94dnr1cnF5iW9ueCV4eW0AAAAAABQZegA9OLoaNjOMTzU0kpg4N6XfPz+o/T9Asv8/QLf+P0C3/j9Atv4/QLP+QECz/0BAsspBQqtQQUKrCYJ3fwBOTKICTkyiAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wCOjdsBjo/ZAo2K3wCLeP84c2f0pWle7/9rXPL/aFrx/2ha7/9nWu3/Y1no/2Rb4P9aUtL3T0rFsEFBtFBFRbIcNTuLBf//5ALLxOAA/vT6FP3z+YT78PXj+e3x//fs7//47fD/+O3x//fs7//26u7/9ejt//Pl6//z5ev/8OPs/+7h7P/u4Oz/693q/+ja6f/l1uj/4dHo/9rO5v/Txej/yLnv/6mb6v94btX/W1XK/0dFvP9TUKb/cGuX/39+m/+EgZH/hYGP/4eDjf+KiIv/ioiL/4qIi/+KiIv/ioiL/4qIi/+KiIv/ioiL/4iGiv+HhIn+ioiL/4eEif6Fgoj+hYKI/4OAh/hycYDDbW1+PnVujgAmMn"
    Sec08_HeaderLogo_9 = s
End Function

'--- Sec08_HeaderLogo_10 : Dashboard header, toolbar buttons (Layout/Update/KPI/Chart/Slicer/Save/Reset/Deploy) and the embedded company logo bitmap (chunk 10 of 18)
Private Function Sec08_HeaderLogo_10() As String
    Dim s As String
    s = s & "8JNTShJjw3pW46NZjFODWa+zk5qv88Pq/9Pj+x/j8/sv8/P7L/QECz/z5Asf0/QLL/QUGy3kRDsX9EQ7APU0z/AEJCjwJFRZIAaGa0AGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAkI/bAJCP2wGQktcCjorgAY1z/F54aPXHaV7v/2hb7v9nWfD/Z1nw/2ZZ7v9kWOr/Z1vs/2FW6P5YUd77UU3R20xKvIZHRqstR0efCv///wD47vse/PL7sfvw9vL67vL/+Ozw//fr7//47fH/+Ozv//bq7v/15+3/8+Xr//Pl6//w4+z/7uHs/+7h7P/r3er/6Nnp/+TV6P/g0ej/2s7m/9PF5//Iuu//qZzr/3hu1f9bVsr/T0u+/ldTrv5saaD/f3+b/4SBkv+FgY//h4OO/4qIi/+KiIv/ioiL/4qIi/+KiIv/ioiL/4qIi/+KiIv/ioiL/4qIi/+KiIv/hoSJ/oWCiP6EgYf/gHyI/HBzj+Rpbo9RAAIrCS4zh0I2NZ6qPDeX5zo4n/w4OKj/ODip/zg4qf86Oqz/Ozqu/jw8sP4/P7P/PkCx/T9As/89Pa7qNjWinDg2pBMyMp0AQ0SCAHBuvABoZrQAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGhmtACxrtsAy8fmAMrN2gGUkd8Ci3T9gXpq9uZoXu//aFvu/2dZ8P9nWfD/Zlnu/2RY6v9nW+7/Z1ru/2NY6f9gV+P8ZFzb5VZQyL5STcR6SknOMtnT/DD98vnO/PD2+/nu8v/47fH/+O3w//jt8f/47fH/9+vv//Xn7f/z5ev/8+Xr//Dj7P/u4ez/7uHs/+vd6v/o2en/5NXo/+DR6P/azub/0sXn/8y99P+toPD/eG7V/1tVyv9PS7/+V1Ov/mtoof9/f5z/hIGS/4WBj/+Hg47/ioiL/4qIi/+KiIv/ioiL/4qIi/+KiIv/ioiL/4qIi/+KiIv/ioiL/4qIi/+Hhoj/hoOH/oWBif6Df47/fHmY+VZdfKcvNXOUOTSOyTs3lu86OKH+ODin/zg4qf84OKn/ODip/zg4qf85Oav/Ozqu/jw7sP08O7D9PTuz/zw6sfo6OKrCR0O+I////wA9K/8AaGezAWhnswBoZ7MAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AGFfsQCMiMAA3ei4AYiJyQWAbfuUdWf29mle7/9oW+7/Z1nw/2da8P9nWu7/Zlrs/2VZ6/9lWev/ZFnr/2RZ6f9mXOb/YVnY/FlUzdlgWcyuoJPip/Pm+e788Pf/+u7y//jt8f/47fH/+O3x//js8P/36u7/9eft//Pl6//z5ev/8OPs/+7h7P/u4Oz/693q/+jZ6f/k1ej/4NHo/9rO5v/Sxef/zb/1/7Cj9P"
    s = s & "97cNr/W1XK/09Lvv5XU67+bGmh/39/nP+EgZL/hYGP/4eDjv+KiIv/ioiL/4qIi/+KiIv/ioiL/4qIi/+KiIv/ioiL/4qIi/+KiIv/ioiL/4iHh/+GhIf+hYKK/oWAkP9/epr/Yl2M6EQ/h+k7NpT/Ojif/zg4p/84OKr/ODiq/zk5q/85Oav/OTmr/zk5q/85Oav/Ojqu/js6r/08OrP+PDqy/zo6rNo6O6g9Oj2hCAkRhwBlYb8BZWG/AGVhvwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAaWe0AU5SjAB/etECf3jaEXhp96RxZPT/aV7v/2hb7v9nWu//aFrv/2hb7v9oW+7/ZVnr/2RY6/9kWOv/ZFjr/2NY6f9pX+j/X1jY/2xk0/+jluX/7N/7//vv9//67vL/+e3y//jt8f/47fH/+Ozw//fr7v/16O3/9Obs//Pl6//x5Ov/7+Ls/+7h7P/r3er/6Nnp/+TV6P/g0ej/2s7m/9LF5//Nv/X/saP1/31y3/9fWNH/U07E/1lUsv9raKH/f3+c/4SBkv+FgY//h4OO/4qIi/+KiIv/ioiL/4qIi/+KiIv/ioiL/4qIi/+KiIv/i4iL/4uJjP+KiYv/iYeI/4eFh/+Gg4n/hICQ/396mv5sYpb/UUiS/zw4mf84OKj/ODip/zg4qf84OKr/OTms/zk5rP85Oaz/OTmq/zg4qf85Oaz/Ojmt/jw6sv48OrT/Ozqv5Tc3pVs0NJ0SGBqZAFFPqAJRT6gAUU+oAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGZktgBmZLYBAAAAAIFz6ASAcuokeWr1tXJl8/9pXu//aFvu/2hb7v9oW+7/aFvu/2hb7v9oW+7/Z1rv/2dZ8P9mWe7/ZFnq/2hd7v9jWuL/fHLi/7qu9P/t5P3/+O71//ru8v/67/P/+O3x//jt8f/47fH/+Ozw//fq7v/16O3/8+br//Pl6//y5Ov/7uHs/+vd6v/o2en/5NXo/+DR6P/azub/0sXn/82+9f+xo/X/f3Tg/2Ze2/9aVND/XFe3/2toof9/f5z/hIGR/4WBj/+HhI7/i4iL/4qIi/+KiIv/i4iM/4uIi/+LiIv/i4iL/4qIi/+MiY3/jouO/4uMiv+KiYv/iYeI/4eFiP+FgY//gHuY/m1pkv9STY//OzeY/zg4qP84OKn/ODip/zg4qf84OKn/ODip/zg4qf84OKn/ODip/zg4qv84OKn/Ozmw/z06tf89OrTpPDiwczw2rho0KNkAQkKPA0JCjwBCQo8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYFy9AGBcvQEABQAAiXH7Bohw+zJ7ave8cmXz/2le7/9oW+7/aFvu/2hb7v9oW+7/aFvu/2hb7v9nWu//Z1nw/2"
    s = s & "dZ7/9nW+7/aFvu/2JZ4v98c+L/vK/0//Pn/f/67/X/+e7y//rv8//47fH/+u7z//nu8v/47PD/9+vu//Xo7f/z5uv/8+Xr//Lk6//u4ez/693q/+jZ6f/k1ej/4NHo/9rO5v/Sxef/zL71/7Sm9f+IeuD/aF/b/1pU0P9cV7j/a2ih/39/nP+EgZL/hYGP/4iEj/+OiY7/i4iL/42Jjf+Pio//joqP/46Kj/+Oio//jouO/4+Nj/+OjY7/i4yK/4uMiv+JiYj/h4aG/4eFiP+FgY//dXCS/1dRk/87N5j/ODio/zg4qf84OKn/ODip/zg4qf84OKn/ODip/zg4qf84OKr/Ojqs/zo6rP87Oq7+Ozqv/zo6ruk8OK9zPDavGjQo2gBCQo8DQkKPAEJCjwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74AX1u+AQAAAACGdf0HhXT9QXtr98JxZPP/aFzu/2hb7v9oW+7/aFvu/2hb7v9oW+7/aFvu/2hb7v9oWu//aFvv/2le7v9qX+j/Ylne/4F35f/Iuvr/9en+//nu8//57fH/+u/z//jt8f/77/T/+e7y//js8P/36u7/9ejt//Pm6//z5ev/8uTr/+7h7P/r3er/6Nnp/+TV6P/g0ej/2s7m/9LF5//MvvX/tab0/4p84P9sY9//W1XR/1xXt/9raKH/f3+c/4KBmf+DgJP/iYSP/46Kjv+LiIv/jYmO/4+Kj/+Pio//j4qP/4+Kj/+PjI//j42P/46Njv+LjIr/i4yK/4qLif+J"
    s = s & "iYj/h4aH/4WCjv97d5H/XFeS/zs3mP84OKj/ODip/zg4qf84OKn/ODip/zg4qf84OKn/ODip/zg4qv86Oq3/Ojqt/zs6rv47Oq7/Ojqt6Tw4r3M8NrAaNCjaAEJCjwNCQo8AQkKPAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF9bvgBfW74BAAAAAIZ3/AeGdvw7emv3wHFk8/9oXO7/aFvu/2hb7v9oW+7/aFvu/2hb7v9qW/D/aVvv/2hc7v9pXu//aV/u/2df4f9kW9n/iH7l/9HE/v/26v//+e70//nt8f/67/P/+O3x//rv8//57vL/+Ozw//fq7v/16O3/8+br//Pl6//y5Ov/7uHs/+vd6v/o2en/5NXo/+DR6P/azub/0sXn/8y+9f+1pvT/inzh/25l4v9cVdL/W1e3/2toof9/f5z/gICc/4KAlf+JhI//j4qP/42Jjf+Oi47/j4uP/4+Lj/+Pi4//j4uP/4+Mj/+PjY//j42P/42MjP+LjIr/i4yK/4qLif+Hhof/hYKL/3x6j/9fXZL/PjyZ/zk4qP84OKn/ODip/zg4qf84OKn/ODip/zg4qf84OKn/ODiq/zo6rf86Oq3/Ojqv/js6r/86Oq3pPDivdT02sBs1Kd0AQkKOA0JCjgBCQo4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYFy9AGBcvQETJ0YAinr6BIl4+id7bPi3cmX1/2le8P9pW+//aFvu/2hb7v9oW+7/aFvv/2pb8v9pW+//aFzu/2le7/9oXe7/ZFvi/2he2f+SiOb/183+//bq/v/57vX/+e7y//rv8//57vL/++/0//nu8v/47PD/9+ru//Xo7f/z5uv/8+Xr//Lk6//u4ez/7N7r/+ja6f/l1uj/4dLo/9rO5v/Uxuj/zb/1/7Wm9P+KfOH/bmTh/15X1f9dWLn/a2ih/35/n/+AgJ3/goCU/4mEjv+Pio//j4uP/4+Mj/+PjY//j42P/4+Nj/+PjY//j42P/4+Nj/+PjY//j42P/42MjP+LjIv/iouJ/4iGh/+Fgon+fnuQ/2Nhlf9CQpv/OTmo/zg4qf84OKn/ODip/zg4qf84OKn/ODiq/zg4qv85Oav/Ojqt/zo6rf88O6z+PDus/zo6reg6N6xrOjarFywkyABGRZUDRkWVAEZFlQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABnZbUAZ2W1ARENAgCKivcCh4b4GHxv/a51Zvn/bFzz/2xc8/9pW+//aFvu/2ha7/9nWfD/aFnw/2hb7/9oW+7/aFvv/2dZ7/9jWOr/b2Lj/6CU7P/i2v//9uj+//ns9f/67vL/++/0//vv9P/77/T/+e7y//js8P/36u7/9ejt//Pm6//z5ev/8uTr/+7h7P/u4ez/6dvp/+bY"
    s = s & "6P/k1uj/287n/9rJ7P/RwPj/tKb0/4p84f9vZeL/ZFze/2Ndwf9raaT/fn+o/4CAn/+DgZT/iYWP/4+Kj/+PjY//j42P/4+Nj/+PjY//j42P/4+Nj/+PjY//j42P/4+Nj/+PjZD/j4uP/42LjP+Li4r/iYeL/4aCjv+CfZf+Z2SZ/kNCm/85Oaj/ODip/zg4qf84OKn/ODip/zg4qf85Oaz/Ojqt/zo6rf86Oq3/Ojqt/z49p/8/Pqf/Ozqu4DU3oE0wM5ANCBF3AF1ZuQJdWbkAXVm5AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGlnswBpZ7MA////ANPa9AGxrvcJh3X+lHpo+fRsXPP/bFzz/2tc8v9rXPL/a1zy/2tb8v9rW/L/aV3w/2lf7/9oXe//Z1nw/2NY6/91Zuf/qpvv/+jf/v/26P7/+u32//vv8//77/T/++/0//vv9P/57vL/+Ozw//fq7v/16O3/8+br//Pl6//y5Ov/7uHs/+7h7P/s3ev/6Nrp/+XX6P/g0ej/3Mrt/9HA+P+0pvT/i33g/3Vp4f9nXt//ZmDG/3Fvq/9+f6n/gICf/4eFmP+NiZP/j4qP/4+Nj/+PjY//j42P/4+NkP+RjpH/kY6R/5GOkf+RjpH/kI2S/46Llv+PipH/jYuM/4uLiv+Kh4v/hoKO/4J9l/5rZ53+TEmh/0A/pP85Oaj/OTmr/zk5q/84OKn/ODip/zk5rP86Oq3/Ojqt/zo6rf87Oq7+Ozus/zs7q/87OazQNzibLBwxNwIAAA"
    s = s & "AAZmO5AWZjuQBmY7kAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AF5buQBEQLUAcHGvAfTn7wKNdf1ye2n42Wxe9P9sXPP/bFzz/2xc8/9sXPP/bFzz/2xc8/9pXfD/aV3w/2pc8f9nWe//Y1nk/3xt4/+0ou//7uH///fo/v/87vf//e/0//3v9P/97/T//e/0//ru8v/47PD/9+ru//Xo7f/z5uv/8+Xr//Lk6//u4ez/7N7r/+zd6v/p2+n/5dfo/+LS6P/cyu3/0MD4/7an9f+SgOT/fG3n/21j3v9rZMb/c3Ct/39/qf+DgqD/ioea/46Klf+Pio//j42P/4+Nj/+PjY//j42Q/5GOkv+RjpL/kY6S/5GOkv+RjpP/j4yU/4+Kkf+Ni4z/i4uK/4qHi/+Ggo//gn+X/3Jvn/9XU6H/QT2b/zk4pv84OKv/OTmq/zg4qf84OKn/OTms/zo6rf86Oq3/Ojqt/zs6r/07OrH/Ojmu9Dg2pbI3N5wZMyvFAI+H4wBpZ7IAaWiyAGlosgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAdHLLAHVyywFwcMUCNjG/AItz/E57avi5cGL1/2xd8/9sXPP/bFzz"
    s = s & "/2xc8/9sXPP/bFzz/2pb8f9qXPH/alzy/2dZ7/9jWt//gnPi/76r8v/15v//+On+//zt9//+7/T//u/0//7v9P/+7/T/+u7y//fs8P/36u7/9ejt//Pm6//z5ev/8uTr/+7h7P/s3uv/693q/+na6f/l1+j/4tLo/9zK7f/QwPj/uKj2/5eC5/9/b+r/cWXh/21mx/9zcK3/f3+p/4aDoP+LiJr/jouV/5CMkP+QjZD/kI2Q/5CNkP+QjZH/kY6S/5GOkv+RjpL/kY6S/5GOkv+RjpL/kIyQ/42MjP+Li4r/ioiL/4eEjf+Dgpf/eHai/2Bco/8/Opb/Ozmi/zo6qf85Oar/OTmq/zk5qv86Oqz/Ojqt/zo6rf86Oq3/Ozqv/Tw6s/87OLHgNjSjhDY0pA8dG6gATUyZAVZVowBoZrQAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACOjdsAjo3bAY+O2QKOjN8AiHr+J35v+ZN3Z/b9bF3z/2xc8/9sXPP/a1zz/2xc8/9sXPP/bFzz/2td8f9pXfD/Zlnu/2Na3/+Heef/xrb3//vv///56v7//O33//7v9P/+7/T//u/0//7v9P/77vP/+O3w//fr7v/16e3/9Obs//Pm6//y5Oz/7uHs/+7h7P/s3uv/6drp/+XX6P/i0uj/3Mrt/9DA+P+5qPf/mITo/39v6/9wZOf/bGXL/3Nwrf9/f6n/hYOg/4uHmv+PjJb/kY6S/5GOkv+RjpL/kY6S/5GOkv+RjpL/kY6S/5GOkv"
    s = s & "+RjpL/kY6S/5GOkv+QjZD/jY2M/4uLiv+KiYv/i4mM/4WElv97eab/ZWGq/0I9l/9APp7/PT2o/zo6rf86Oq3/Ojqt/zo6rf86Oq3/Ojqt/zo6rf87Oq/9Pjy0/z88tsU8N65MPDetCVJbWwFERJECRESRAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI+O2wCPjtwAi4vYAH570wD//+8Ci3f9YX9t+eNrXfD9a1zy/2pb8v9pWvH/bFzz/2xc8/9pW/D/alzx/2pb8f9kWen/Ylne/42C6P/Owfb//PH9//nq/v/87ff//u/0//7v9P/+7/T//u/0//7v9P/77vP/+Ozw//fr7v/16O3/9ejt//Pm7f/u4ez/7uHs/+ze6//p2un/5dfo/+LS6P/cyu3/0MD3/7us+f+fjPD/gnHx/3Zo6f9waMv/cnCt/39/qf+Fg6D/i4ea/4+Mlv+RjpL/kY6S/5GOkv+RjpL/ko+T/5SRlP+UkZT/lJGU/5SRlP+UkZT/lJGU/5KPkv+OjY7/i4yK/4yMi/+QjZD/i4iZ/356pv9nY6r/SEaZ/0JAn/89Paj/Ojqt/zo6rf86Oq3/Ojqt/zo6rf86Oq3/Ozqv/Ts6r/1HQ7r1T0rAmzY4kBMyNYYCeHS6AVxbqAFcWqgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjo3bAHFuyQBEP7AAX1u+AXFW/wCMdf44gW36qHRl9PlwY/b/al3z/2hZ8f9sXPP/al7x/2pd8f9sXPL/a1zy/2da7P9iWd7/k4fo/9PI9//69P7/+u7///3u9//+7/T//u/0//7v9P/+7/T//u/0//zu8//47fH/9+vu//bp7f/05+z/8+Xr//Hk6//v4ez/7N7r/+na6f/l1+j/4tLo/9zK7f/QwPf/u6z6/6CN8v+DcfL/d2np/3Fpy/9ycK3/f3+p/4WDoP+Lh5r/j4yW/5OQk/+TkJP/k5CT/5OQk/+UkZT/lZOV/5WTlf+Vk5X/lZOV/5WTlf+Vk5X/lJGU/5CPkf+Ojo7/jo2O/5GOkv+LiJr/g3+g/29rnv5JSJj/QUGf/z09qP86Oq3/Ojqt/zo6rf86Oq3/Ojqt/zo6rf87Oq/9Pj2x/0JAtcJHQrlYLjxXAz9BhANLSpoB////AHVywAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX1u+AF9bvgBfW74Blnz/AIhu/RGDbvtcf2755nRn+f5rYPP/aFzw/2pe8f9pX+//a13y/2xc8/9rXPL/aFvt/2dd3f+dkur/2tL8//n1/v/78Pr//e/2//7v9P/+7/T//u/0//7v9P/+7/T//O7z//jt8P/36+7/9unt//"
    s = s & "Tm7P/z5ev/8+Xr/+/i7P/s3uv/6drp/+XX6P/i0uj/3Mrt/9C/+P+7rPr/oI3x/4Jx8v92aen/cWjL/3Jwrf9/f6n/hYOg/4qHmv+QjZb/lJGU/5SRlP+UkZT/lJGU/5SSlf+VlJb/lZSW/5WUlv+VlJb/lZSW/5WUlv+Vk5X/kpGS/4+PkP+QjpH/kI2R/4qHmv+FgqH/dnOg/1VNl/9FQp//PTyo/zo6rf86Oq3/Ojqt/zo6rf86Oq3/Ozqu/j09sf1AP7TwPz2zfjYwqRkTMUEAQ0OOAkBAjQFKSZYAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74AX1u+AF9bvgBoYMwAinT/AIt1/h+Hc/2ld2j58W1i9P9pX/D/aV/v/2pe8P9rXfL/bFzz/2tc8v9pXe3/bGHb/6Sa7P/g2P//+fX9//vw9f/97/T//u/0//7v9P/+7/T//u/0//7v9P/87vP/+O3w//fr7v/26e3/9Obs//Pl6//z5ev/7+Ls/+ze6//p2un/5dfo/+LS6P/cy+v/0cL2/72u+v+gjfH/gnHy/3Zp6f9xaMv/cnCt/39/qv+Fg6D/ioea/5CNlv+UkZT/lJGU/5SRlP+UkZT/lJKU/5WUlv+VlJb/lZSW/5aUl/+VlJb/lZSW/5WUlv+SkpP/j4+Q/5GOkv+PjY//i4iV/4WEoP95d6P/YVaU/0hEnv88PKn/Ojqu/jo6rf86Oq3/Ojqt/zs6rv4+PbH/Q0K0+0RBtrJE"
    s = s & "QLc8Pjm1AGBcwAAPFzUASUibAEJCjwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74AX1u+AF9bvgH/vv8Ajnr8CYZz+jp5avfHc2X2/2xh9P9qXvD/a1zy/2xc8/9sXPP/a13z/2lf7v9sYtz/pZvt/+HZ///59f3/+/D1//3v9P/+7/T//u/0//7v9P/+7/T//u/0//zu8//47fH/9+vv//bp7f/05uz/8+Xr//Lk6//v4uz/7N7r/+na6f/m1+j/4tPo/93P5v/UyfP/wbT6/6GO8f+EcvL/eGnp/3Jpy/9zcK3/f4Cp/4eEoP+LiJr/kI2W/5SRlP+UkZT/lJGU/5SRlP+UkpX/lZSW/5WUlv+WlZf+l5WY/peVmP6VlJb/lZSW/5KSk/+Pj5D/kY6S/4+NkP+Mioz/h4aS/3t6mP9iXpD/SEad/z08qv47OrD9Ozqv/To6rf87Oq7+Pj2x/0RDtf5NSrniUk67T19ZvwpMRrsAX1u+Al9bvgBfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF9bvgBfW74ARk"
    s = s & "CuAIaIzwGSjOAChXT6DIBx+lV8cP3Ddmn6/2tc8v9oWvH/aFnw/2pd8/9sYPT/amDu/29k4/+vpvH/6uT///r1/v/98fj///D1///x9f//8fX///H1//7w9//+7/X//e7z//vr8v/76/H/9+nu//Tm7P/y5Ov/7+Ls/+7h7P/s3uv/69zq/+na6f/l1+j/4tPn/9fL9P/Eufr/q5vy/4528/2Cb+/+e2/S/3x1rv+Gg6H/lIqg/5KLl/+RjZL/lZOW/5WTlv+WlJf+l5SX/paVl/6Wlpf/lpWX/peVmP6XlZj+l5WY/pWUlv+VlJb/kpKT/4+PkP+RjpL/kY6S/4yKjf+KiIz/g4CP/2plk/9KSJ7/PTut/z08tP8/PrL+Ozuv/j09sf5FQ7X/TEq53FBTwV9VWMQV36CtAF5avgFfW74AX1u+AF9bvgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX1u+AF9bvgB4dc0AkY7gAJKbxAGNkN4CiXz+EYRz/2d6a/vMbmH2/2pb8v9nWfD/aVvy/2td8/9pXe//dGfk/7Wr8//s5///+vX+//7x+f/+8Pf//vD2//7w9f/+8Pf//vD4//3u9f/97PL//Ovy//zr8v/25+3/8+br//Lk6//u4ez/7uHs/+ze6//r3Or/6tzq/+bY6P/k1Oj/3Mzv/8q89P+snvL/h3Tu/ntr6v92bM//eHOr/4eDnv+XjJ//lIuV/5KOkf+alpn/m5aZ/5uW"
    s = s & "mv+blpr/m5ea/5mZmf+ZmJn/l5eY/5aWl/+Wlpf/lZSW/5WUlv+SkpP/j4+Q/5GOkv+Tj5P/joyO/4uLiv+GhYv/cW2S/0pGl/87OKb+PDuz/j48tP89OrP/RkO6/E1LvsdRUb9rVF/MFkpbtgM5LmMBZF/HAF9bvgBfW74AX1u+AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wCOjdsAjozfAI6O1wKNdf8Ai3D7Hn1s+ml0Z/nFcWH29mtc8v9qWvL/alvy/2la8P95auf/ua70/+3n///68/7//vD5//7w9//+7/b//u/2//7w+P/+8Pj//e30//vq9P/76vb//Ovy//Xn7f/y5Ov/8OPs/+7h7P/u4ez/7N7r/+rb6v/p2un/5tjo/+HS5//dzu7/zb/0/6ud8f+CdOz/dmnp/3Npzv93cq7/iYWn/5eLof+Wi5n/l4+Y/5+ZoP+fmaD/n5mg/5+ZoP+fmZ//nZue/5ybnP+amZn/mZeZ/5iXmP+WlZb/lZSW/5OSlP+RkJL/ko+T/5SSlf+Ojo7/jI2K/4mJjv93dJT/TUmU/zw6pf48PLX/PTy2/z88tvBLRr64VlPIYFRm0RQ9Pn0BOjV0AkJDkQFMS6AAX1u+AF9bvgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wBpZsQAT0q0AF9bvwGZg/8Ah3r+EoFy/FF4aPm3cmH15G5e9PhrXPP/aVvw/3pq7f++sff/8+j///zv/f/+8Pn//u/1//7w9v/+8Pj//vD4//7w+P/97fT/++v2//nq+v/76/P/9eft//Hj6//u4ez/7uHs/+7g7P/s3uv/6drp/+bY6P/l2Oj/3dDl/9vR7//OxPf/rqDx/4p67P+BcOz/fnLS/4J8tP+SjrX/mpCn/5yQo/+elaT/op2l/6Kdpf+inaX/op2l/6Kdpf+inaT/n56f/5+anf+fmJ3/nZib/5eXmP+Wlpf/lZOW/5SRlf+UkpX/lZSW/4+Pj/+Ojo3/jY2T/315mv5UUpn/Q0Or/kFCvPVAQbrdQj+3qUpFuEZtZ8MLaGRTAENCjwNBQo8BQkKPAEJCjwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjo3bAGhlxABRTLUAX1u/AFJSqQFzaN4Cgnj6C4B4+SiHcPp2e2n5yW5h9vtpW/D8e2vs/8Oz9//56v3//vD4//7w+P/+8Pj//vD4//7w+P/+8Pj//vD4"
    s = s & "//7v9f/97fP//Ovz//vr8v/15+3/8+Xr//Hk6//u4ez/7N3q/+vc6v/p2un/5tjo/+XY5//e0t//39Lj/9bH7f+5qvH/noro/5mE5/+Th9H/lJC2/6agtf+oobT/qKG0/6eir/+mpKb/pqSl/6akpf+mpKX/pqSl/6Wjpf+joaL/op2j/6Ccov+dnJz/m5qa/5uamv+Zl5n/lpSX/pWUlv+VlJb/kY+R/5GQkf+RkJf/f36c/1tZoP1PS7PvUUzFvVJPx2lKR78cQz6vBX570AKNjNkCWVmmAAwNWgBCQo8AQkKPAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX1u+AF9bvgBfW74AV1DAAHh7rgGar8MBlaD2BJF3/BiDbvxFdGT2hm5g7syLeOj6z771//3x/P/87/X//vH2//7x9v/+8vf//vL2//7x9//+8Pj//fD2//ru8v/46+//9+nu//Tm7P/z5ev/8uTr/+7h6//t3+X/7N3o/+ja5v/l2OD/5Njg/97S3f/Zzd7/08jh/8zA5P/EtOf/uq7g/7St1v+1sMv/u7W9/7m0u/+3srn/tbC1/7GusP+vrK3/rKmq/6uoqf+qp6j/pqSl/6WjpP+koaP/oZ6g/56dnf+bmpr/m5qa/5mYmf+Wlpf/lZSW/5WUlv+RkZL/j4+O/4yMkP+Egp37XVuqwE1Jt29bVckyXFfOFv///w"
    s = s & "A9OL8AbmvIAZGQ3gEAAAAAX1+sAEJCjwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF9bvgB6eM8Ak5PgAI2L1wHMy/8Alnn/AKJ8/wZ6ZfYkdGbsWKWS6Zfi0vjd/O/7//vv9P/+8fX//vL2//7y9v/+8vb//vL2///x9v/88PX/+e3y//fr7v/26u3/9efs//Tm7P/y5Ov/8OLo/+7g4//s3uj/6Nrl/+TY3P/k2N3/3tLd/9fM2P/UydT/1MrV/9LH2P/KwtL/xcDO/8S+yv/DvcD/v7q9/724uv+6tbf/tbK1/7Owsf+wrK3/rqqr/6uoqf+mpKX/pqSl/6Sio/+gnp//nJub/5uamv+bmpr/mJmY/5aWl/+VlJb/lZSW/5GSkv+Njoz/i4uN/4mGm/dmZrGDPkPKIy8qxwB2bP8AgXf/AGBcwwFbV7sBhILVAI6N2wCOjdsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    Sec08_HeaderLogo_10 = s
End Function

'--- Sec08_HeaderLogo_11 : Dashboard header, toolbar buttons (Layout/Update/KPI/Chart/Slicer/Save/Reset/Deploy) and the embedded company logo bitmap (chunk 11 of 18)
Private Function Sec08_HeaderLogo_11() As String
    Dim s As String
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wCBf80AAAAAAGdltAFdWLwBrra/AW1X9AD///8A6Nr0PvTo/br67vv//PD1//7y9v/+8vb//vL2//7y9v/+8vb//vH1//zw9P/57fL/9uru//bq7v/16O3/9Obs//Pk6f/x4eX/7uDi/+zd5f/o2uP/5Njd/+PX2//d0dr/2M3T/9bLzf/Vy83/0MjK/83FyP/JwsT/xL/B/8K8v/++ur3/vLe6/7m1t/+1srT/sq+w/7Gsr/+vqqz/q6ip/6akpf+mpKX/pKKj/6Cenv+bmpr/m5qa/5uamv+YmZj/lpaX/5WUlv+VlJb/kpKT/46Pjv+NjY3/jYqV9nt7pXFbYcAZgorQAFNToANTU6AAV1aqAF9bvgBfW74AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGhmtABpZ7MAaGa0AP70/wDc1PMD3NX0A+LU8gbh0vFF9O78vfv0+//+8vb//vL2//7y9v/+8vb//vL2//7y9v/88PT/++/0//ns8f/16O3/9ejt//Tm7P/z5ev/8eTo/+3g4v/t4OL/693f/+fa3f/j19z/4NTX"
    s = s & "/9vQ1P/YzdD/1svO/9XLzv/Qx8r/zMTH/8jBxP/Ev8H/wry//766vf+8t7r/ubW3/7WytP+yr7D/sayv/6+qrP+rqKn/pqSl/6akpf+koqP/oJ6e/5uamv+bmpr/m5qa/5iZmP+Wlpf/lZSW/5WUlv+VkpX/kpGS/4+PkP+Pio/4iISUj4J+mCuBe3wAjIvZA4yL2QCMi9kAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AGlnswBoZrQAn5zfAKSh4gJXXNkB6d/qB+rf6lX57fPF/vL2//7y9v/+8vb//vL2//7y9v/+8vb//vL2//vw9P/77/T/+ezx//Xo7f/16O3/9Obs//Pl6//x5Oj/7eDi/+3g4v/o29//5Njd/+LW2//b0NT/2s/T/9jN0P/Wy87/1cvO/9DHyv/HwMP/xL7B/8S/wf/CvL//vrq9/7y3uv+5tbf/tbK0/7KvsP+xrK//r6qs/6uoqf+mpKX/pqSl/6Sio/+gnp7/m5qa/5uamv+bmpr/mJiY/5aXl/+Wlpf/lZSW/5WUlv+SkpP/j4+Q/4+Nj/mSkZWalZaaMpaXgQCPjtwDj47cAI+O3AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADPx+8Az8fvArCr6wL37fQI9+70Yv3y9cz/8vb///H1///x9f//8fX///H1///x9f//8fX//e/0//3v9P/67PH/9ejt//Xo7f/15+3/9OXq//Li5v/t4OL/6Nzf/+jb3v/m2dz/39TX/9jN0P/az9P/183R/9bMz//YztH/0cnM/8W/wv+/ubz/vrm8/8C7vf+/urz/urS3/7Svsf+wrK7/r6yt/7Grrv+wqqz/ramq/6qnqP+npab/pKGj/6Cen/+dnJz/nZyc/5ybm/+YmJj/lpeX/5aXl/+Wlpf/lZWW/5STlf+RkJL/kI6Q+ZaWl6ecnZ06nqCIAI6N2wOOjdsAjo3bAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAOXb9gDl2/YD1870A/3z+Qn99Plm//T4z//z9v//8fX///H2///x9f//8fX///H1///x9f/+7/T//u/0//vs8f/16O3/9ejt//Xn7f/z5Oj/7+Dj/+vd4P/l2d3/49bY/+HU2P/f09n/29DT/9DHyv+9tbr/r6et/6yip/+km6L/mpGZ/5aNlf+Xjpf/mpKY/52Zm/+fm57/o5+i/6mmp/+rp6j/ramq/6yoqv+qp6j/qqeo/6elpv+lo6T/o6Ch/6Cenv+enZ3/nZyc/5qamv+YmJj/mJiY/5aXl/+Wlpf/lZOW/5OQk/+Qj5D6l5iXt5yenESeoJIAhoXTAoaF0wCGhdMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA2NDyANjQ8gO+uO0D/PL6Cf3z+mj/9PjQ//P2///x9f/+8Pf///H2///x9f//8fX///H1//7v9P/+7/T/+uzx//Xn7f/16O3/8+Xq/+/i5v/s3uD/6Nrd/+PX3f/az9X/1MnS/8/Ezv/DuL7/pZ2h/4J9hf9pZG3/XVde/llQWv9VS1n/VkxZ/1hPWv9ZUFr/Xlxj/m1qcf6AfYH+j42P/5aUlf+dm5z/oqCh/6WjpP+lo6T/paOk/6ajpP+loqP/oqCh/5"
    s = s & "+env+enZ3/nZyc/5uamv+bmpr/mJiY/5aXl/+VlZb/k5GU/5GRkfyVlpXMmJqXUZqckABmZLIBZWOxAGZksgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACamN4AmpjeAQAAdAH98vsJ//P8c//0+Nb/8vb//u/0///x9f//8fX///H1///x9f//8fX//u/0//3t8//56e//8+br//Xo7P/v4uX/7eDi/+ve4P/k193/3tLc/9TK2/+8tMX/lY+Z/21jb/9RSFL/PjtG/zMzPv8uLDT/Liw0/zQyPv8zMTz/Lisz/y4rNf8vLEH/PTdH/0tETv9UT1r/YmNm/3h2eP+NjIz/nJub/6KfoP+joKH/o6Ch/6Ogof+joKH/o6Ch/5+env+enZ3/nZyc/5uamv+bmpr/mJiY/5aWl/+Vk5X/lJGU/ZOSk+WSkpRhlZZxAHV0wQDy8/8AaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAL674gC/u+IB////AP/z9wr/9PeB//T33v/y9v/+7/T///H1///x9f//8fX///H1///x9f/+7/T//e3z//jo7//z5uv/9ejs/+/i5f/s3uD/6Nrc/+HV2f/e0t3/qqS7/3Nti/9ORmL/NzJI/zszTf9CPl//SUlv/1NQdf9aVXn/XVl9/1hVef9ST3T/UU5z/0Y/ZP9BN1f/OTJL/zEuQf86OkT/SUVM/2VhZf+CgoP+lJSV/56cnf+fnp7/n56e/6Cenv+joKH/oqCg/6Cenv+enZ3/nZyc/5uamv+cmJv/mpeZ/5aWl/+Vk5b9k5KT55GRlGXb4gAAiYfSAHh2wwBoZrQAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA///vAP//7wD/8/gA/vT1C//09oX/8/bg//H1//7v9P/+8PT///H1///x9f//8fX///H1//3u8//97PP/+env//Pl6//05uv/7+Ll/+3f4f/p3N7/2M"
    s = s & "3Q/6uhqf9mYXb/PjlY/zwyWP9FO2v/XVGL/2lgqf9xabz/e2/D/39yxf98ccP/enDC/3luwf91bL3/cmS1/2NXn/9JQoP/Lyxj/yknPf8yLzn/QD5F/1ZXXP94eXr/lZSV/52bnP+dnJz/n52d/6Ogof+joKH/op+g/6Cen/+enZ3/nZyc/52am/+amJn/lpaX/5WUlv6SkZP3kZCUc4+t7QOFiq0AZGG0AGloswBpaLMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA7ujrAP/0+QD98/cM/vP4jf/y9uT/8PX//u/0//7v9P//8PX///H1///x9f//8fX//Ozy//zr8v/56e//8+Xr//Pl6//u4eT/7uHj/9/T1v+up6r/ZF9m/jo2S/88Nln/WE6B/25fq/9+cMr/f3Pb/35x4/+Dc+b/iHbr/4Bz5v+CdOj/h3fs/4N16P+Hdu3/fXDb/2lgwf9SS6H/PTpl/yswSP8mKjz/NjVC/1hZYv+HhIr+mZeZ/52cnP+gnp//o6Ch/6Ogof+joKH/op+g/56dnf+enZ3/nJub/5mZmf+Wl5f/lZWW/5KPk/+QjJiDgYDxCnp4wQBhXrgBZGG5AGRhuQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADWyv8A/PL6APzy+g388vqf/vH37v/w9f/+7/T//u/0///w9f//8fX///H1///x9f/97PP//Ovy//np7//z5ev/8+Xr//Dj5v/o29//urK2/2Vlaf47OUX/Rj5g/2Zblf+BdcP/h3nY/oFz4P93a93/cWbX/3No2v+Acef/em3k/3xu5f+Dc+r/gXLp/31u6f+AcuT/hHff/4J10/5pY63/NEB4/yAqUP8rK0T/QUFV/3FteP+PjZD/npyd/6Ogof+joKH/o6Ch/6Ogof+in6D/np2d/56dnf+cm5v/mZmZ/5aXl/+Wlpf/kpCT/46KnouBdfYWd3DQAVtYuwJfW74BX1u+AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOHY9gD68PQA+vD0Dfrw9Kn98fXz//D1//7v9P/+7/T///D1///x9f//8fX///H1//3s8//86/L/+env//Pl6//z5ev/9+ru/8nAxP96dnz/OzpF/0Y6Uv9qXJb/iHfU/5B86/+Ed+H/inzg/4x+4P+JfN3/h3nZ/oh62/6LfN//jX7h/4x+4P+Je9z+iHrc/ox93v+FeOH/eGvh/3lt2/9qY7//PESE/xwrTf80NUH/UFFX/3t5ff+ZlJr/op2k/6Ogof+joKH/o6Ch/6KfoP+enZ3/np2d/5ybm/+ZmZn/lpeX/5aVl/6RkZL/k5OWlqSksCKNjLABV1PBAmBcvgFgXL4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA49v0AP/88QD16/MR+vDzq/3x9PT/8PX//u/0///w9f//8fX///H1///x9f//8fX//ezz//zr8v/56e//8+Xr//Lk6v/l2t7/nZGe/1FFXP8+N1X/XlaI/3luwv+CdOP/gXHn/4x+4v+4rfD/zMP2/9DH9f/OxPP/u7Dx/8C29P/KwPX/zMP0/8m/8v/LwvL/w7n0/6OW8P98bev/emzw/4Jy5f9oYLn/Pz94"
    s = s & "/ywrP/82N0D/aGZu/5KMlP+gmp//o5+h/6Ogof+joKH/op+g/56dnf+enZ3/nJub/5mZmf+Wlpf/lpWY/pGRkv+NjJKpiIOaOoeEqAGCgd4Cg4HUAYOB1AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wD//6sA///7APXr9hP88vau/vL29f7x9f/+7/T///H1///x9f//8fX///H1//3w9P/77fL/+uvx//fp7v/05ez/8ePp/7+4u/9uZ3L/PjNQ/1JGe/9yarz/fnTb/3tu4/97buD/koXd/9PL+P/y6v//+PD///Tv///Z0P7/4Nn//+7p///18P//9O////fw///m3f//tqr1/39v6v95avX/gnHw/31v2f9iWKP/OTNU/y0rPv9YUmD/iIGK/6CboP+jn6H/o6Ch/6Ogof+in6D/oJ6f/56dnf+dnJz/m5qa/5iYmP+Yl5j/kZKS/42NkreLiJhKjIqoApCP6AKPjtwBj47cAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7ujrAPv44QD///8A9+34FP/0+bX+8/f5/vH1//7v9P//8fX///H1///x9f/+8fX//O/0//jt8v/37O//9unt//Tm7P/u4OX/mJSW/1BOVP9AOlj/ZVij/31w2/9/ceb/fG7j/4F03/+UiNj/2NH1//jw///89f//9vP//9/Y/P/l3v7/8u3///r4///8/P///fb9/+vi/P+3rfL/fG7l/3hp9P95avH/fW/n/3Nowf9RSHj/LytI/0Y/T/96cXv/op2k/6Ogof+koaL/pKGi/6Shov+joKH/oJ6e/56dnf+dnJz/nJub/5uamv+SkpP/kpOUu5iZmlCVlqgCjIvoAo6N2wGOjdsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA5dz2AP///gD37fgV/vP5vf7z9/3+8fX//u/0///x9f//8fX///H1///x9f/+8fX/+u7y//fs7//36u7/"
    s = s & "9Ofr/+TW2f9/fH7+TUVU/1RHdf9uZMD/fm/m/31u7f99b+j/g3ff/5WJ2P/c1PX/+vf///39///5+P//6eL+/+jh/v/x7P7/+vj///z7///6+P//6eP+/7et8v99b+X/eGn0/3hp9/97bO7/eG3T/2Jdnv80MVH/NS88/2ZeZf+emZ3/o6Ch/6WjpP+mpKX/paOk/6Ogof+joKH/oJ6f/56dnf+dnJz/m5qa/5WTlf+Vk5a7l5mZUJWVqAKMi+gCjo3bAY6N2wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wD57/UA4NzhAfzx9x3+8/m//vP3/f7x9f/+7/T///H1///x9f//8fX///H1///x9f/+8PT/+e3w//js8P/w5en/xb3A/2ZgZf9EQVj/XliV/4Jz3/98bPL/eWn1/3xs8P+DdeX/lYnY/9zU9f/7+P///f3///r5///t5///59zy/+/m9P/6+P///Pv///r4///p4/7/uK/x/39y5P95avT/eGn3/3ps7f+Act//fHDF/kc+af8sJzf/UE1Q/5GPkP+joaL/pKKj/6Wio/+mpKX/paOk/6Ogof+ioKH/oJ6f/52cnP+bmpr/lpWX/5WVl7uXmJlPlZaoAo2L6AKOjdsBjo3bAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wDu6OsA7ujrAPvx9QD17esB/fL3H/7z+cD+8/f9/vH1//7v9P//8fX///H1///x9f//8fX///H1///z9v/67/L/+u/z/+3i5/+spqr/W1NZ/0lEX/9pYqX/iHfp/3dp9v9zZvj/d2nz/4Bx6P+Thd7/2dH3//n3///9/v//+vr//+3o///i1u3/6uHw//v5///8/P//+vj//+nj/v+5r/H/gHTk/3lq9P94aff/eWvv/4Fy5/+Gdtv/XVCP/zMtRP9FREX/hoSE/6ajpP+loqP/pKKj/6akpf+mpKX/paOk/6Shov+hn5//np2d/5ycnP+Xl5j/l5iYwpqbm1mTlKQCbWnbAnd0zQF3dM0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7ujrAO7o6wDu6OsA//b6APPs6wH88vcf/vP5wP7z9/3+8fX//u/0///x9f//8fX///H1///x9f//8fX///T3//vw9P/88PX/6t/k/52Wm/9WT1X/UUtl/3Jpq/+Iduj/dWj2/29j+P90ZvP/f2/q/5GC4//Tyvj/9PD///r4///39P//6eP//9zS8v/m3vX/+PX///n2///49P//6eL+/7mv8f9+ceT/eWr0/3hp9/95avL/gnHv/4t35v9qWab/OTJP/0NBQv+Df3//qqen/6ekpf+koqP/pqSl/6akpf+mpKX/pKKj/6KfoP+fnp7/np2d/5iYmP+YmZjPm5ybaJSUngI4Mc4BXlq6AV5augAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA7ujrAO7o6wEAAAAA8+vqAf3y9x/+8/nA/vP3/v7x9f/+7/T///H1///x9f//8fX///H1///x9f//8/b//fH1//3x9v/n3eH/lZCV/1RNU/9VTmf/d22u/4h36P91aPb/bmL3/3Vm9v+BcO//kIHi/8rB9//o4f//7+j9/+zl/P/h2v7/29P9/+Tb+//w5/v/7uf7/+7n/v/i2v3/ta"
    s = s & "rx/31u5v94afT/eGn3/3pq9f+FcvP/kHrr/2pZpf84Mk7/Q0FC/4SBgP+tqqv/qKWm/6Sio/+mpKX/pqSl/6akpf+koaL/o6Ch/6KfoP+enZ3/mJmZ/5iZmNmZm5h1lpiZAgAA/wBoZrQAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wDu6OsA7ujrAQAAAADw6eoB/fL3H//0+cP+8/f//vH1//7v9P//8fX///H1///x9f//8fX///H1///x9f//8fX///L2/+fd4f+Xk5f/VE5U/1ROZ/93ba7/iHfo/3Vo9v9zZvv/fGz+/4Rz9v+Nft7+wLX1/9jQ/v/c0/X/3dPs/9/Y/P/j2///5dn1/+PU6v/d0ur/3NT4/87D+P+pm/D/f3Dt/3lq9f93afv/emv4/4Nx8f+IduP/aFij/zkyTv9DQUL/hIGA/62qq/+qp6j/paOk/6akpf+mpKX/pqSl/6Shov+joKH/op+g/56dnf+bmpr/mpqZ2pmbmHaYmpoCoaSSAGdlsgBoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7ujrAO7o6wDu6OsBAAAAAPDo6gH98vcf//T5xP/z9//+8PX//u/0///x9f//8fX///H1///x9f//8fX//fD0//7w9///8/v/7eHo/5+Znv9XT1X/VE5n/3dtrv+Id+j/dWj2/3dr/f98b/r/gHDt/5CC4f/Yy/f/8er//+/r/P/w6vj/8uj+/+DV7v/k3O3/9O/4//Dq+P/v6/3/49z9/7er9v+AcO//eGr6/3dp/f96a/j/gnHx/4d14v9oWKP/OjJY/0RBS/+DgH//rKiq/6qnqf+mpKX/pqSl/6akpf+mpKX/pKGi/6Ogof+in6D/np2d/56cnP+YmJjrk5STi5KStASUlW8AionYAGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA7ujrAO7o6wEAAAAA8OnqAf3y9x//9PnE//L3///w9P/+7/T///H1///x9f//8f"
    s = s & "X///H1///x9f/87/T//fD4///1/v/t5e3/o52k/1hQV/9UTWf/d22v/4h36P94avb/eGv6/3ps9P99buj/koPi/+TX+P//+v//+/r///r3///89v//49jq/+jh7v/9/v//+vj///z6///v6///vrT3/35v6v94afj/d2n9/3pq+f+Cce//h3bc/2lZo/86M13/RUNR/4aDg/+rpqj/qKWm/6ekpf+opqf/pqSl/6akpf+koaL/o6Ch/6KfoP+enZ3/np2d/5iXl/ORkZKXiYm9CAAA/wCGhc4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wDu6OsA7ujrAQAAAADw6eoB/fL3H//0+cT/8vf//vD1//7v9P//8fb///H2///x9v//8fb//vH2//zv9P/98Pf///X+/+zp8v+vqbH/XVZe/1NNaP9za6z/hXXk/3xs8/94afb/eWrx/35v5v+Sg+L/4tT4///5///7+///+vj///v6///o3e7/6+Hx//r4///6+P//+/v//+3p//+6sfL/fG/i/3pq8/95avr/e2v1/4Fx6P+Ddc/+YVOT/zYvVf9JR1b/jYuM/6qmp/+opKX/qaan/6yoqf+opab/qKWm/6Wio/+joKH/op+g/56dnf+enZ3/mJeX"
    s = s & "8pGRk5eGhb8Lbmv/AYOCygAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7ujrAO7o6wDu6OsBAAAAAPDp6gH98vcf//T5xP7y+f/+8Pb//u/1//7w9//+8Pj//vD4//7w+P/+8Pf//O/1//3w9//99Pz/8+/2/9DIzP9xanL+U05q/mlioP98cdf/fm7u/3pq9v96a+7/f3Hh/5KD4f/f0ff//ff///z8///7+///+fb//+Pa7f/j2fD/9fD///v7///6+P//6uT+/7eu7v98cNr/fm3v/35t8/9+b+z/gXPd/3xvvv9LQGz/LihE/1NQX/+alpn/q6eo/62pqv+tqar/ramq/62pqv+tqar/qaan/6WjpP+ioKD/np2d/56dnf+Yl5f0kZGTmYaFvgtraf8Ag4LKAGhmtABoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA7ujrAO7o6wEAAAAA8OnqAf3y9x//9PnE/vL5//7v9v/+7/X//vD4//7w+P/+8Pj//fD3//zv9f/98Pf//fL6//zy+v/77/X/8eLo/5WIkfZZTGD5XU9+/39yyv99bej/fGzy/31t7P9/cOD/joDc/tfN9v/38///+vj///r2///68f//5Njt/+DX8P/w7P//+/v///r4///t5///urLw/31x2f+AcOv/fG3t/4l36v+Rf9v/d22w/zo1V/84M0X/amVu/6Ogof+sqKn/ramq/62pqv+tqar/ramq/62pqv+tqar/qaan/6Sio/+enZ3/np2d/5qYmv+VlJish4a/C1hY/wCCgcoAaGa0AGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wDu6OsA7ujrAQAAAADw6OoB/fL3H//0+cT+8vn//u/2//7v9f/+8Pj//vD4//7w+P/+8Pj//fD3//7w+P/88/z/+/L6//3v9f/68/j/t7C0+21gb/xSQWr/cGGt/4R13f+Bcu"
    Sec08_HeaderLogo_11 = s
End Function

'--- Sec08_HeaderLogo_12 : Dashboard header, toolbar buttons (Layout/Update/KPI/Chart/Slicer/Save/Reset/Deploy) and the embedded company logo bitmap (chunk 12 of 18)
Private Function Sec08_HeaderLogo_12() As String
    Dim s As String
    s = s & "n/fG3m/39x4f+Qgt//2dH3//n0///7+f///Pr///r1///q4fj/5t35//Dr///6+P//+/j//+7p//+8tPP/fXDg/39v7v9+bu3/iXnp/4Z3y/5dVIb/NDFH/1BNWf+Jg4v/r6qu/62pqv+tqar/ramq/62pqv+tqar/ramq/6qnqP+npab/paOk/6GfoP+enZ3/nZyc/52ana6Sl6wRhpK8AQAAzABoZrQAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7ujrAO7o6wDu6OsBAAAAAPHq6gH98vcf//T5wv7y+f/+8Pb//u/1//7w+P/+8Pj//fD3//3w9v/98Pf//vD4//zz/P/78vr//O70//rz+P/b0tX/joKM/1NFZP9cT4v/fXLE/4V43f+Bc+P/gHLg/5CC3//Wzff/8Or///Ht///08f//9fD//+be///f1///5d////Hs///y7v//5t7//7aq9f96bOT/fGzu/4Ry7/+Edt7/bGKq/j84W/86Okb/cG50/6Sfpf+2r7T/rqmr/62pqv+tqar/ramq/62pqv+tqar/qqeo/6mmp/+opab/pKKj/6Gfn/+hnp//oZ6gq5adpRSOm6gCIxHLAGhmtAFoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA7ujrAO7o6wEAAAAA8+zrAfzy9x/+8/nA/vH5/f7w9v/+7/X//vD4//7w+P/98Pb/++/1//zv9f/+8Pf//vL6//3x+f/87/T//Ozz//rp7v+6rrL/bWV0/lBHbv9kX5v/f3XH/4h73v+Dd+D/in3c/rqv8P/Kwff/ysH4/8/G+v/Uyvn/x7z4/7+0+P/Dufj/zsb6/8/H+v/Fufr/oZPx/3dq5f99bur/hHTj/3Npuv9NSnn/Ly1D/1JUWf+Qj5D/tK6z/7OstP+vqq3/rqqr/66qq/+tqar/ramq/62pqv+tqar/ramq/6yoqf+npab/pKGi/6KfoP+gnp+rlpylFI2bqgIgDssAaGa0AWhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAO7o6wDu6OsA7ujrAf///wDz7OsB/PL3H/7z+cD+8fn9/vD2//7v9f/+8Pj//vD4//7w+P/+8Pj//vD4//7w+P/+8Pj//vD4//7w+P/+8Pf//u/1/+nc4P+qoKb/VE1e/09Hbf9kXpr/fXXG/4t84P+DduD/hHfj/4V35P+Fd+T/hXjl/4Z45f+HeuT/h3vk/4d85P+IfeT/iX3k/4Z55P+Cc+f/gnLq/4d25f9yabv/TUt0/zExPv85OUT/fH1//6ikp/+yq7b/rKW1/7SutP+zr7H/sa2v/7Cqrf+tqar/ramq/62pqv+tqar/ramq/6yoqf+loqP/op+g/6Cen6uWnKUUjZuqAiAOywBoZrQBaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7ujrAO7o6wDu6OsA//T4APTt6wH98vcf/vP5wP7x+f3+8Pb//u/1//7w+P/+8Pj//fD2//zv9v/+8Pj//vD4//7w+P/+8Pj//vD4//7w+P/46O3/+Ofs/93R1v+Yk5r/VlBh/05GbP9oX5v/g3jJ/4d63P5/cOL/fW7n/35v6v9+bun/emvp/3ps6P9+b+n/f3Dp/3tt6P98buP/g3Pl/4l45v+Fd9z+cGO//1RKhP84NE7/OztD/25tcP+loqT/"
    s = s & "tK6x/7Krsv+zrLX/tq+1/7avtP+zr7L/sa2v/7Cqrf+wqq3/rqmr/62pqv+tqar/ramq/6Wio/+hnp//oJ6fq5acpRSNm6oCIA7LAGhmtAFoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADu6OsA7ujrAO7o6wD06vUA////APrw9xn+9Pm5//L5+/7w9//+7/X//vD4//7w+P/98Pf//fD2//7w+P/+8Pj//vD4//7w+P/+8Pj//vD5//fo7f/76/D/+ezx/9rT1v+QiZD/WFFk/1BFaP9jV4v/a2Sq/3pwxv+Edtz/iHjm/oZ35f+Acuj/fW/p/39w6f+CdOj/g3Xk/4Jz2v+AdNP/eW3A/2Vamv9JQ3L/QjtY/01IV/9vbXL/oZ2f/7Svsf+4srX/t7G1/7avtf+1sLX/trG1/7WwtP+zrrL/sq2v/7Grrv+wqq3/rqmr/62pqv+tqar/qaan/6Shov+gnZ+rlp2lFI6cqQIjEcsAaGa0AWhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA7ujrAODX9AD///wA+O75Ef70+qz/8vn1//D3//7v9f/+8Pj//vD4//7w+P/+8Pj//vD4//7w+P/+8Pj//vD4//7w+P//8vn/++7y//vu8f/+8fX/+O3x/9PL0P+TjJP/ZVtn/1NIXv9MS27/YVqO/3Jor/96b8L+fXDH/npwz/93a9P/d2zS/3twzv99b8X+dWix/2hin/9WUoD/RT1Y/z09Rv9ZVl3/f3yB/6Sho/+7trn/tK+y/7m0t/+8t7v/uLK3/7eztv+1srX/tbG1/7WwtP+zr7H/sq2v/7Gsr/+wq63/rqqr/66qq/+uqav/qKWl/KOhoaaUmasRhZGzAQAAywBoZrQAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADv6uoA8ezpAP3z+wD88voN/PL7nP7y9+z/8PX//u/1//7w+P/+8Pj//vD4//7w+P/+8Pj//vD4//7w+P/+8Pj//vH4///09///"
    s = s & "8vb/++7y//fq7v/15+z/9urv/+Xa3f+1rK7/e3Z4/1tVYf9LQmD/TEhv/1VQg/9fUZH/Y1yb/2FYof9hWaH/YVqa/15Qjv9TR3X/SENh/0VCU/9QSlH/Z2dq/5WUlv+wra//uLS3/724u/+3tLf/u7e6/766vf+9ubz/u7a5/7e0tv+2s7X/trG1/7WvtP+zr7H/sq+w/7Ktr/+xq67/sKut/66pq/+tqqj1rKqnm42MwQs6PP8AhobLAGhmtABoZrQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wDu6OsA//j8AP709wv+9PeJ/vP24v7x9f/+7/X//vD4//7w+P/+8Pj//vD4//7w+P//8fn///H6///x+v//8fn///T3///y9v/+8PT//O3y//rq8f/26e7/+Ovv//Hi5v/Xys3/rKaq/4aBj/9pYXT/V0pj/1A/Yf9QQWL/UEFj/0s+ZP9HPWP/UEBh/1pMY/9dU2P/cWtz/5SSk/+lo6T/ura4/7+7vv+9ubz/wbu+/8C7vv/Bu77/wby//8G7vv+9ubz/vLi7/7u3uf+5tLf/tbK1/7WytP+zsLL/sq2v/7Crrf+uqav/ramq/62qqPWtq6eYmJbDBu7rUACIh9QAaGa0AGhmtAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA5dz1AOXc9QLGv/EC/vP4Cf70+HH+8/fV/fH1//zv9f/+8Pj//vD4//7w+P/+8Pj//vD4///x+v//8fr///H6///y+v//9Pf///L1///w9f/+7vT//Ovy//bo7v/36e7/+Orv//Tm6//l297/1cvS/8C1vP+onKT/kIiS/4N8hf+CeoP/fnaC/3x0gv+JgYv/mpKa/6mepf+5r7P/yMDD/8vDxf/Kw8b/xsDD/8K+wP/EvsH/wry//8K8v//CvL//wry//8C7vv++ur3/vbm7/7u3uv+4tLf/tbK1/7SxtP+zsLL/sa2u/62pqv+tqar/p6Wk5KChnYKdnbADoaKbAI+O3QCOjdsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADLxO4Ay8TuAq2o6wL16/MI9uzzYfzx9sz+8vb//fD1//7w+P/+8Pj//vD4//7w+P/+8Pj///H6///x+v//8fr///H6///z+P//8vb///D0//7u9P/76/L/9uju//Tn7P/z5ev/9ebq//fp7P/56+//8+Xp/+XZ2//Vy83/ycHC/8W9vv/Durz/xLu9/8rCw//QyMr/2M7R/9zR1f/c0NT/2MzQ/9DHyv/LxMf/ycPF/8fBxP/Gv8P/xL7B/8K9wP/CvL//wry//8C7vv++ur3/vbm7/7u3uf+4tLf/uLS2/7aytf+zrrH/r6yt/6+srf+lpKPSl5mXbJiamwKCfv8Bjo3bAI6N2wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJOS2wCTktsBAArEAevh6Qbs4ulL+/D0wP/z9v//8fb//vD4//7w+P/+8Pj//vD4//7w+P//8fr///H6///x+v//8fr///H6///x9v/+7/T//e3z//vr8v/26O7/9uju//bn7P/04+j/8+Pn//Lk6f/z4+j/8+Pn//Pi5v/y4uX/7d7g/+vd3//r3uD/6Nrc/+LW2P/c0dX/2s/T/9nO0f/Wy8"
    s = s & "7/0snM/9DHyv/Oxsn/zMTI/8zEyP/Jw8X/xcDC/8K9v//DvcD/w73A/8C7vv++ur3/vrm8/7y3uf+7trj/ubO3/7WvtP+zr7H/s7Cx/6inqMaYmZpdmJipAouJ6QKOjdsBjo3bAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA3djoAN3Y6AG/wNwA9OryAvPq8Rn+8fWk//L1//7y9v/+8Pj//vD4//7w+P/+8Pj//vD4///x+v//8fr///H6///x+v//8fr//vD2//3t8//86/L/++vy//bp7v/56vD//Ovy//rp8P/z5ev/7+Hk//Hj5//y4+j/7eDi/+vc3v/q293/5tjb/+HV2P/h1dj/4NTX/93R1f/ZztL/18zP/9bLzv/Wy87/08nM/8/Hyv/Px8r/z8fK/8/Hyv/LxMf/xcDC/8XAwv/FwML/w77A/8K8v//BvL//vrm8/7y3uv+5tbf/tbG1/7awtf+1sbX/qqiqupWWl0+WlqgCjozpApCP2wGQj9sAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD/+u4A//juAAAA/wD47fgB5t31Bv/0+Iv+8vf3/PD1//7w9//+8Pj//vD5///x+v//8fr///H6///x+v/+8vv//vL7///x+v/97fT//Ovy//zr8v/86/L/++rx//zr8v/66vH/9uju//Pl6//y4ub/8uLm//Pi5//x4eX/7d/h/+vc3v/p2tz/5tnc/+PX2//g1dj/39PW/97T1v/d0dT/2c7S/9fMz//Vys3/08nM/8/Hyv/Px8r/z8fK/87Hyf/Mxcj/ysPG/8rCxv/HwcT/xL/B/8S/wf/CvL//wbu+/765vP+7t7r/urS3/7aytf+vra+gm52dLqamrQHLxvUBwr7kAcK+5AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAO7o6wCkot8AgoLYAJKR3QL///8A//X7Z/70+tT88fn//PD7//7w+f/+8fn//vL7///x+v"
    s = s & "//8fr///H6///x+v/+8vv///H6//7u9//97fT//e3z//3t8//97fP//ezy//vq8f/46e//9Ofs//Pk6f/z4+n/8+Po//Pi5v/x4eX/7d/h/+vc3v/q293/59rd/+PX2//h1dj/4NTX/+DU1//d0tX/2c7S/9jN0f/WzM//z8fK/83FyP/Oxsn/z8fK/8/Gyv/Lw8f/y8PH/8jCxf/Fv8L/w73A/7+6vf+/ur3/wLu+/8C7vv+/ur3+vLnA8rm4v3mpqagNvLq6AP///wD//+8A///vAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wCOjdsBjYzbA////wD67fxD/O/8rv3x/f/78P///fD6//7x+f/+8vv///L7///x+v//8fr///H6///y+v//8fr//vD5//7w9//+8PX//vD0//7w9P/97vP//ezy//vr8f/46e7/9Ofs//Tm7P/z5er/8+Pn//Pi5v/x4eT/7d/h/+vc3v/q293/59rd/+PX2v/h1dn/4dXZ/+HV2f/e0tX/29DU/9jO0f/Sycz/z8bJ/9DHyv/QyMr/zsbJ/8zEyP/Lw8f/yMLE/8S/wf/DvsD/xb/C/8S+wf/DvcD/wr2//7+6vf2qqrnmnqG3ZPn/vQB1dbIA"
    s = s & "urbSAO7o6wDu6OsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjo3bAI6N2wGOjdsD//juAM7A4S/k1vOX8uT9//ns///97/v//vD5//7x+//98vz//vD7///x+v/+8vv//fL8//7x+//+8Pn//vD5//7w+P//8fb///H2//7w9v/97fX//Ovz//vr8//36fD/9eju//Tn7f/z5O3/8uLo//Pi5//w4eX/7N7i/+rb3//q297/5tne/+TY3v/k197/5Nfe/+HV2v/d0db/2s/U/9rP1f/XzNH/18zR/9TK0P/QyM3/0MfN/83EzP/Iwcj/xL/E/8bAxf/Pxcr/yL/H/7u2vv+xrbX/pqSt/nZ4me5eYo5tAAA/AjQ7eAA8Q4EAjo3bAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACOjdsAjo3bAY6N2wOKgHgAgn2SN6mezqDKvPb/7d3+//fo/v/46f3/+On9//jq/v/46v7/+u3+//vv/v/77///++/+//vv/v/77/7/++/+//vv/v/77/7/++/+//rs/f/46f3/+On9//jp/P/26fP/9Oby//Hh+v/v4e7/7+Hr/+3f6v/r3On/6tzo/+fZ5//m2Of/5Nbn/+LT5//i0+f/3dDl/9vP3//bz93/2s7j/9rP3P/azt//2M3j/9bL4f/Rx9v/z8Ph/8e93P+/utT/vrjR/6+pvP+Vk7H/gH2l/3Bslf9bW4b+RUKL9EA7kW8AAAAAenrCAHJzuACOjdsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHJvygByb8oBcm/KApKAyAB9cr4jkIHei52M6/+zoO3/vavz/8S19v/Lvfb/0sT3/9vP/f/d0///4Nf+/+bd/v/q4v//7uX///Dm///w5v//8Ob///Dm///05///9+f///bm///25v//9ub///bm/f/y5fz/7eL//+zi+//o3/r/5Nr5/+HX9//h"
    s = s & "1/j/4Nb4/9/W9//f1ff/3tP4/9fL+P/Nv/P/ybvv/8i67//As+//tKfe/6ujzv+koML/m5i+/5CLvv+Kg7r/fni1/3JusP9raar/Yl+m/1dWo/9UUqb/VE+q/0xJpPpIRKmzSEKuQwAAcQCPjtsCj47cAI+O3AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAWla7AFZRuABfW74Bho3CAYx4/wmIdPRbinXu2pB46/qUfe/+moXy/p+N8P6lle7/rJ7y/66i8f+zp/D/vK7x/7+y8//DuO7/xLrr/8S57P/Euez/xLns/8i77P/LvO3/y7zt/8u87f/LvO3/y7zt/8i67P/Bt+v/wbfs/7617f+3rej/sKbj/7Cm4v+uptz/rqba/6uj3P+lnt7/oJjf/5mP2/+UidX/kYXR/4t/z/+BdcT/enK1/3FtqP9lZKX/XVmv/1hUq/5TT63/Tkqv/0lGrP9FQ7b/RUO4/0pHvv5PS8j/UE3J8VVSz2xfW9oXUEvWAIaF0wKGhdMAhoXTAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74AfnzRAJCP3ACGhdYCimv/AIlx+yiKcfiQjHH08I1x9P+Ocvb9jnT0/Y147f+MeO3/jXvn/5KA5f+Ygun/mILo/5iH4P+ZiNz/mYjd/5mI3f+ZiN3/l4fb/5WH2v+Wh9r/lofa/5aH2v+Vhtr/kYPa/4t/2f+Lftf/i3/X/4d50v+Cc8z/gHHM/3xww/96cL7/dWvA/2xjxf9tY8X/bWTG/2pfv/9lWrb/ZVq1/2Vbtf9mW7b/YFi0/1ZRs/9TTsD+Uk3B/lZRyP9aVc7/WlTP/1VQ1/5TT9f+Uk7W/1JO1flRT9TRUlDWOwAAAAC4vbMAZWOxAWRhrwBlY7EAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABjX8AAY1/AAGJfwAKpif8AeXPkCoZx"
    s = s & "7kmMcfPhkHX5/5R5/f+We/7/lnz+/5Z8/v+Wfvj/l3/z/5d/9P+Xf/T/l3/z/5eA8/+Xf/P/l3/z/5d/8/+QfOz/jHro/4x66f+Meun/jHrp/4d26/+Ccuz/gnLr/39x5P9+cOP/hHTl/4h36f+Ccev/gHHr/35w5f98buT/emzq/3pr6v96a+n/emvp/3dq5/9zZ+H/cmfg/3Nn4P9uZOD/Z17e/2Ja2v9hWtr/Ylra/2Ja2/9hWdv/WVTg/1RQ2v5STtb/UU3S50tJvYlJSLceGRicAI6N2gJ3dcMAz9D/AGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFtXvABbV7sAVVC4AIuL1gH/AP8AiHH0H4tx+Z+Nc/3rkXX+/5R4/f2Wev//l3z//5Z7/v+Vefz/lHj6/5R5+/+We/z/l339/5d9/f+WfP3/kHr8/414+/+MePr/jHj6/4x4+v+NePr/inb7/4d1/P+HdPv/g3L6/4Jy+f+Ec/r/hXP7/4Nw+/99bvr/e235/3pt+f96bPn/eWr1/3lq9P92aPX/cmb1/29j8v9sYe7/aF7q/2RZ5/9fVuX/W1Xh/1pV4P9dVeL/XlXj/1pV4P9YU+H/V1Lg/1VR3P9ST9"
    s = s & "fLUE3JNk1KuAcdGMUAeHfFAZKR3wB1c8EAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX1u+AGBcvwBsacQAg4HQAY1o/wCKdv0JiXb7SYpz/r6Odf35k3f8/pR3/f2Vef3/lXn//5V5//+Vev7/lXr+/5d7//+XfP//lXv//5F5//+NeP7/jXj//414//+NeP//i3f//4t3//+Kdv//iXb//4l2//+Fc///hHP//4Ry//+Ecf//hHD//4Fv//+Ab///fm///3tv/v95bPv/eGr3/3Vn9/9vY/f/a2Dy/2lf8P9nW+3/ZFjq/2BV5/9cVOT/W1Pk/1tT5P9aU+P/V1Lh/1VQ2/9VUNr/Uk/U505MyJNVUdcQ////AIqKywFQTZsAXlyqAF5cqgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYV+tAGFfrQBcWacAb2agAIib5wKFh/gQh3P9go51/N+Td/z/k3b8+5N3/P6UeP3/lXn+/5d7//+XfP//l3z//5d8//+Ue///jXj+/414/v+NeP7/jXj+/4x4/v+Jdv//iXb//4l2//+Jdv//iHb//4Vz//+Ec///hHL//4Rw//+EcP//hHD//4Vw//+Ccf//fXH//3pt/f95avj/dWf3/29j9v9pX+//aV/v/2ld7/9mWuz/Ylfp/19U5v9eVOb/W1Pj/1hT4f9XUuD/U07W/lJO1f9QTc+9SUi0SUlIswX///8AiIbWAYB/zQBoZrQAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAaGa0AI6N2wCOjdsBkZfRAodz+AKFbvxAjnT8rZJ2/P6Sdvz8k3f8/pN3/P+UeP3/lXr//5d8//+XfP//l3z//5R7//+OeP7/jXj+/414/v+NeP7/jHj+/4l2//+Jdv//iXb//4l2//"
    s = s & "+Idv//hXP//4Rz//+Ecv//hHD//4Rw//+EcP//hHD//4Fw//98cP7/em39/3lq+P91Z/f/b2P2/2lf7/9pX+//aV3v/2Za7P9iV+j/XlTm/11U5v9aU+P/WFPh/1dT4P9UT9j+UlDX71NR14dcVd4LU0/CAmRkmwFqaLYAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGhmtABoZrQAionXAI+O3ACPlc8AgWzwAHpf+QKJc/1tinP95o50/vySdvz8lHj+/pV5//+Vev//l3z//5d8//+XfP//lnv//5R5//+OeP7/jXj+/414/v+MeP7/iXb//4l2//+Jdv//iXb//4h2//+Fc///hHP//4Ry//+EcP//hHD//4Rw//+EcP//gW/+/3pt/f96bf3/eWr4/3Vn9/9vY/b/aV/v/2lf7/9pXe//Zlrs/2JX6P9eVOb/WlPj/1hT4f9YU+H/WFPh/1hT4P5XUteqVlLLQUpCqwBDQ5IDRESQASkqdgBmZLIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    Sec08_HeaderLogo_12 = s
End Function

'--- Sec08_HeaderLogo_13 : Dashboard header, toolbar buttons (Layout/Update/KPI/Chart/Slicer/Save/Reset/Deploy) and the embedded company logo bitmap (chunk 13 of 18)
Private Function Sec08_HeaderLogo_13() As String
    Dim s As String
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACOjdsAcG3JAFVPugBdWrwBjHP/AIdx/S2Kcv6TjXP//JJ2/P+Ud/z/lHj+/5V5//+Wev//l3v//5d8//+XfP//lnv//5N5//+SeP3/kXf8/5B3/f+Md/7/jHf+/4p3//+Jdv//iHb//4Rz//+Ec///hHP//4Ry//+EcP//hHD//4Vw//+Bb/7/em39/3pt/f95avj/dWf3/29j9v9pX+//aV/v/2ld7/9mWuz/Ylfo/15U5v9ZU+L/WFPh/1hT4f9YU+L9XFXn6VxV4mNVU60KamPwAEFBjQFBQY8ASUmWAGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABfW74AXFi8AGFdvwF+l6gBiHD5B4tx/UmMc/7PkHX++ZJ2/P+Ud/z+lHj+/5V5//+We///l3z//5d8//+XfP//lXn//5R4/v+UeP3/k3j9/5B4/f+QeP3/jXf+/4t3//+Kd///h3X//4Vz//+Ec///hHP//4Rx//+Bcf//gXH//35v/v96bf3/eWv9/3Zo+f9zZvj/b2P2/2lf7/9pXe//Z1vt/2VZ6v9iV+n/XlTm/1lT4v9YU+H/WFPh/1hT4fdbV+OcXVnlK1RP3gCOjNwBAAAAAF9erABCQo8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF9bvgB7ec4A////AIqJ2AKJZ/8AiXH5Gopy+niLc/7gj3T+/ZJ2/P2Td/z+lXn//5d7//+XfP//l3z//5d8//+Wev//lXn//5V5//+UeP7/k3f8/5N3/P+QeP3/jXj+/414/v+Kdv//hnT//4Rz//+Ec///hHP//35x//99cf//fHD+/3ps/f93af3/c2b5/3Fl+P"
    s = s & "9vZPb/aV/v/2hc7v9mWez/ZFjq/2JX6f9eVOb/WVPi/1hT4f9YU+H8VlTg1FRZ00ROY7cJAAAcAI6N2wKOjdsAjo3bAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGxquABsargAbGq4AURWdwB8qkIBh3fvHYhz/amLc/7vj3T+/5N3/PyVef//l3v//5d8//+XfP//l3z//5d8//+Wev//lHj+/5N3/P+Td/z/k3f8/5B4/f+NeP7/jXj+/4l2//+FdP//hHP//4Rz//+Ec///fnH//31x//98cP7/emz9/3dp/f9zZvn/cWX4/29j9v9pX+//aFzu/2hb7v9mWuz/Ylfo/15U5v9ZU+L/WFPh/1dU4eBQWd2FT1rcFkJX/gBiXb4BpabqAI+O3ACOjdsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYmCuAGJgrgAVEV8AjIraAZK5jwGHefAHiHX8QIxz/76PdP7/k3f8/pV5//+Xe///l3z//5d8//+XfP//lnr//5V5//+UeP7/k3f8/5N3/P+Td/z/kHj9/414/v+NeP7/iXb//4V0//+Ec///hHP//4Rz//9+cf//fXH//3xw/v96bP3/d2n9/3Nm+v9vY/f/bGH0/2lf7/9oXO7/ZVnr/2RY6f9iV+n/XlTm/1lT4v9aVOTzXFboqVRY3ipQWNsFMEb/AGRitwF6ecoAjIrZAI6N2wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAaGa0AIB/zQCPjtwBkJLWAY2F6gOKefkLiXH+bI"
    s = s & "50/dKTd/z/lHf9/ZV5//6Xe///l3z//5d8//+Vef//lXn//5R4/v+Td/z/k3f8/5N3/P+QeP3/jXj+/414/v+Jdv//iHX//4Z0//+Ec///hHP//35x//99cf//fHD+/3ps/f93af3/c2b6/29j9/9sYPP/aV/v/2hc7v9nWu3/ZVnr/2JX6f9eVOb/WVPi/1pU5MdeV+peZFPjB5aP1QKIitABWFafAGlosgBpaLIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjo3bAI6N2wCOjN4Ajo/WAo5x/wCCbfshjHP8fZF1/OOTdvz7lHf9/pV5/v6Wev7/lnv//5V5//+VeP7/lHj9/5N3/P+Td/z/k3f8/5B4/f+NeP7/jXj+/4l2//+Jdv//h3X//4Rz//+Ec///fnH//3xw//97b/7/emz9/3dp/f9zZvr/b2P3/2xg8/9pX+//aFzu/2hb7v9mWuz/Ylfp/15U5v9YVeHqVlbhg15Y6R1nY9QBl5fWAY+O3QGFhNIAaGa0AGhmtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAjo3bAHNxvwBcWagAaGa0AYxy/wCJcPotjHP7lZF2/PCTdvz9k3b8/ZN3/f+Vef7/lXn//5R4/f+Td/z/k3f8/5N3/P+Td/z/kHj9/414/v+NeP7/iXb//4l2//+Hdf//hHP//4Rz//9+cf//e2/+/3pt/f96bP3/d2n9/3Nm+v9vY/f/bGDz/2lf7/9oXO7/aFvu/2VZ6/9hV+j/XlXn9VVY4KFMW9s5TlnjAGBcvQJcWLsAZ2TCAI6N2wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACOjdsAaWe1AGdlswBoZrQBqYf/AIODuAOIdfU2iXT9xZF2/PeTd/z+k3b8/JN3/f+Vef//lHj9/5N3/P+Td/z/k3f8/5N3/P+QeP3/jXj+/414/v+Jdv//iXb//4d1//+Ec///hHP//35x//97b/7/em39/3ls/f93af3/c2b5/29j9/9sYPP/aV/v/2hc7v9lWev/YVfo/15V5/heVunPW1bkQQBQDAEACP8AZ2S2AWZktQBoZbcAjo3bAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAZWOxAGVisACMitsCeSj/AIB17A2Fc/VUkXb71pN3/P+Td/z9k3b8/JN3/f+Td/z/k3f8/5N3/P+Td/z/k3f8/5J3/P+QeP3/jXj+/4l2//+Jdv//h3X//4Rz//+Ec///fnH//3tv"
    s = s & "/v95bP3/d2n8/3Nm+v9xZfj/b2P3/2xg8/9oXO7/aFvu/2VZ6/9iV+r/YVfq215X31xdVdQRQD7tAIuL2wJwb7oAVlOdAGlnswAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGhmtABzcb8Ac3LAAI6N2wGWl90Bd37KAoBy5BGQdPt0knb81pN3/P+Td/z9k3f8/5N3/P+Td/z/k3f8/5N3/P+Td/z/k3f8/5F3/f+NeP7/iXb//4l2//+Hdf//hHP//4Ry//9/cf//e2/+/3ls/f93af3/dGf7/3Jm+f9vY/f/bF/z/2hb7v9mWuz/Ylfp/2FW6eFhWOuGXljMFz5EMgJmZrUCk5LgAf///wCAfswAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wCOjdsAjo3bAI2M2wGSl9MClIX1AYVt/hePdfyCknb86ZN3/P+Td/z/k3f8/5N3/P+Td/z/k3f8/pN3/P+Td/z/kXf9/414/v+Jdv//iXb//4d0//+Dcf//gW/+/31w/v96bv7/eGr8/3Vo/P91aPz/cmX5/25i9f9rX/L/aFzu/2RZ6v9eVub2W1Xko11X6CxlYOwEOT1YAT8+jAFjYrAAjo3bAI6N2wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAACOjdsAjo3bAI6O2gGOjNwChVj/AIVu/DGMc/yTknf885N3/P6Td/z/k3f8/5N3/P6Tdvz8k3f8/pN3/P+Rd/3/jXj+/4l2//+Jdv//h3P//4Jv//97bf3/em39/3lr/f91aPv/cWX4/3Fl+P9uYvb/a2Dy/2le7/9oXu7/ZVnq+lpZ47pOWtxTh3DtAYyK2gKIh9MB0M7/AEJCjwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wCOjdsAmZfjAIiH1gBFUo0Bb2LuA4px+zqNc/utknf885N3/P+Tdvz9k3f8/5N3/P6Td/z/kHf9/454/v+NeP7/iXb//4l2//+Hc///gm///3pt/f96bf3/eGv9/3Vo+/9wZPj/bmL2/21h9f9qXvP/Zlnu/2RZ6vpoXOzNZV3qVjNYzQdDRP4Ad3XAAXd1xQBrabcAQkKPAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAI6N2wCAftIAq6ztAGJcxAGNc/8AinL6Coty+0COdPutknb87pN3/P+Tdvz9k3f8/5N3/P+Sd/z/kHf9/411//6Jdv//hnX//4Vy//+Cb///em39/3hr/f94av3/dWj7/3Fl+P9tYfX/a2Dy/2ld8f9lWO/9YFfq3Whd6mF3Z+wTX1/oAF9bvgFfXLAAXl2jAF5cqgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF9bvgBfW74AYV28AHJl1ACGav8AlX/5BIlx+0iPdPysknb87pN3/P+Td/z9k3f8/5N3/P+Rdv3/jXP//ot0//+JdP//hnL//4Jv//96bf3/d2r9/3dp/f90Z/v/b2P2/21f9P9pXfH/Zlru/WJX7dheVul8Yl7gFTwo/wB5ddYBXlq+AVpVwgBiX7sAaGa0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABoZrQAaGe0AGhkuQH/4/8Ai3L9AIlw+0mOdPusknb88ZN2/PyTd/z/k3f8/5F2/f6Nc//+jXP//o"
    s = s & "xz//6Icf//gm///3pt/f93av3/d2n9/3Rm+v9tYPX/bF30/2ha8f9jWO3eX1fphltV3xJiYcsBnJzcAY6N2wJzccsAcm/KAF9bvgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGhmtABTTcsAYFy9AV5bvQJgXLwAh2/8AYpx+0KNc/utknf88pN3/P+Td/z9kXX9/I1z//6Nc//+iHH//4Vw//+Cb///em39/3dq/f93af3/c2b6/21h9f9sYPT9ZVvv8mVb7I94beYgbmbiA42N2QKQj9sBjo3bAIqJ2QCKiNgAX1u+AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AGJeuwBdWcAAYV29AWdltQGOcv8AiG/7CYpx+zeMc/yRknb865F1/f+Oc//+jXP//o1z//6Mc//+iHH//4Jv//96bf3/eGr9/3dp/f9zZfr/a1zz/mNZ7uJnXeuAcWXpK6CT3wSTktwCkZHaAYyL2wCOjdsAjo3bAI6N2wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX1u+AGJfuwAoGPsAaGa0AHBpxAGOcv8Admf1BIZv/C+Qdfx+j3T+wY1z/++Nc///jXP//4py//+EcP//fW3+/3hr/f93af3/d2j8/3dm+uRzYvWoYljrZmhg5iLLutcDr82/AI+O2wGNi9wAjo3bAI6N2wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AGhmtABoZrQAaGa0AGtmvAGza/8AekT/AIZu/RiOcv5SjnP+j410/7SMc//Dh3H/zIBv/tB7bP3Qemr9znlq/cJ4aPqve2n5gYpy/TvnvP8DcF/6AHt/ywGOjdwCjo3bAI6N2wCOjdsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaGa0AGhmtABoZrQAlJPhAI2N2gGOjdsDjIHpAaA//wCIhPkKhoD6JYR3/U2Gcfxphm/8c4Zv/HKGb/xuh2/9Sodz8yF+f8oHhm/8AX9u7QBfW70BYFy+ASkinQCOjdsAjo3bAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD/////////AAA//////////////////AAAH/////////////////"
    s = s & "wAAB/////////////////4AAAP////////////////+AAAD/////////////////gAAA/////////////////wAAAH////////////////8AAAB/////////////////AAAAf////////////////wAAAH////////////////4AAAAf///////////////+AAAAD///////////////4AAAAAH//////////////4AAAAAA//////////////+AAAAAAP//////////////AAAAAAB//////////////wAAAAAAf/gP///////wB/4AAAAAAA/gA///////8Af+AAAAAAAP4AP//////8AD8AAAAAAAB4AB//////+AAOAAAAAAAAMAAP//////gAAAAAAAAAAAAAD//////wAAAAAAAAAAAAAAf/////8AAAAAAAAAAAAAAH//////AAAAAAAAAAAAAAB//////wAAAAAAAAAAAAAAf/////8AAAAAAAAAAAAAAB//////AAAAAAAAAAAAAAAf/////gAAAAAAAAAAAAAAH/////4AAAAAAAAAAAAAAB/////+AAAAAAAAAAAAAAAf/////gAAAAAAAAAAAAAAH/////4AAAAAAAAAAAAAAB/////+AAAAAAAAAAAAAAAf/////gAAAAAAAAAAAAAAH/////4AAAAAAAAAAAAAAB/////+AAAAAAAAAAAAAAAf/////gAAAAAAAAAAAAAAH/////4AAAAAAAAAAAAAAB//////AAAAAAAAAAAAAAAf/////wAAAAAAAAAAAAAAf/////8AAAAAAAAAAAAAAH//////AAAAAAAAAAAAAAB//////4AAAAAAAAAAAAAAf/////+AAAAAAAAAAAAAAP//////wAAAAAAAAAAAAAH//////8AAAAAAAAAAAAAB//"
    s = s & "/////AAAAAAAAAAAAAAf//////8AAAAAAAAAAAAAP///////gAAAAAAAAAAAAP///////4AAAAAAAAAAAAD////////AAAAAAAAAAAAB////////8AAAAAAAAAAAA/////////gAAAAAAAAAAA/////////8AAAAAAAAAAA//////////AAAAAAAAAAAP/////////+AAAAAAAAAAD//////////gAAAAAAAAAA//////////4AAAAAAAAAAP/////////+AAAAAAAAAAD//////////gAAAAAAAAAAf/////////4AAAAAAAAAAH/////////+AAAAAAAAAAB//////////wAAAAAAAAAAf/////////8AAAAAAAAAAH/////////+AAAAAAAAAAB//////////gAAAAAAAAAAf/////////4AAAAAAAAAAH/////////+AAAAAAAAAAB//////////gAAAAAAAAAAf/////////gAAAAAAAAAAH/////////4AAAAAAAAAAB/////////+AAAAAAAAAAAf/////////gAAAAAAAAAAH/////////4AAAAAAAAAAB/////////+AAAAAAAAAAA//////////gAAAAAAAAAAP/////////4AAAAAAAAAAA/////////+AAAAAAAAAAAP/////////gAAAAAAAAAAD/////////4AAAAAAAAAAA/////////+AAAAAAAAAAAP/////////gAAAAAAAAAAD/////////4AAAAAAAAAAA/////////+AAAAAAAAAAAP/////////4AAAAAAAAAAD/////////+AAAAAAAAAAA//////////gAAAAAAAAAAP/////////4AAAAAAAAAAH/////////+AAAAAAAAAAB//////////gAAAAAAAAAAf/////////4AAAA"
    s = s & "AAAAAAH/////////+AAAAAAAAAAB//////////gAAAAAAAAAAf/////////8AAAAAAAAAAH//////////AAAAAAAAAAD//////////wAAAAAAAAAA//////////8AAAAAAAAAAP//////////AAAAAAAAAAD//////////wAAAAAAAAAA//////////+AAAAAAAAAAP//////////gAAAAAAAAAD//////////4AAAAAAAAAB///////////AAAAAAAAAAf//////////wAAAAAAAAAf//////////8AAAAAAAAAH///////////wAAAAAAAAB///////////+AAAAAAAAA////////////gAAAAAAAAf///////////8AAAAAAAAH////////////AAAAAAAAB////////////wAAAAAAAA/////////////AAAAAAAAP////////////4AAAAAAAP////////////+AAAAAAAD/////////////wAAAAAAB/////////////8AAAAAAAf/////////////gAAAAAAP/////////////+AAAAAAP//////////////gAAAAAD//////////////8AAAAAB///////////////gAAAAAf//////////////+AAAAAP///////////////gAAAAD///////////////4AAAAB////////////////AAAAB////////////////4AAAA/////////////////AAAAf///////4lQTkcNChoKAAAADUlIRFIAAAEAAAABAAgGAAAAXHKoZgAAAAFvck5UAc+id5oAAIAASURBVHja7P15gC1ZVteLf9beEXHGzLx557p1a+qau6q7q2eauRsRsE"
    Sec08_HeaderLogo_13 = s
End Function

'--- Sec08_HeaderLogo_14 : Dashboard header, toolbar buttons (Layout/Update/KPI/Chart/Slicer/Save/Reset/Deploy) and the embedded company logo bitmap (chunk 14 of 18)
Private Function Sec08_HeaderLogo_14() As String
    Dim s As String
    s = s & "FQVBRU+D0Q0acoKsMDfSoPEVHQp80TcWIQBR4PB0DmQYZukYaGnruri665bt0h5zxDxN57/f7YO+LEOZl3qr5Vt6purqpzMzPOOXEi4sRaew3f9V3Cgbzk5Ke++mv2bDt8+DD9fo/jx48z2dmVIJhJIfrgX/n8oF/7E/iuRTCYC12e+KqPZYd/e2WYb8gRO+W4VHJEvVkhyACVPO1SAUFQkF3BnFdjzlbWnn3iUOf83d/yt3f5438ZTnQZ9yxOlB/4mR/mda99s0zxoqr4aamf9L/9Wf3or/waWVEwOHIYgLv/5t+43pfwQJLI9T6AA7ky2U/p27K9ucX/96P/ka//P/8u+bFl6QQjKzteVra89EqM0SzDmEIz7WoWhkhYJoRVCXoEr4cILBHooZopCigoEn9hLLAOnPeEC1P8WoXfNGLGPelMCumUFIX7T591n/+f3//dvF1OSG4snWGfw7feokduvYVyNGa8tbXnuA+MwfWVAwPwEpHLGYBcRTpq+P3Orn7tzv1QTXnqvtycOO86hmxJbH4UMTchcgIJRxFdBZZRHaDaR0OX4HM0WFUPGkA1PqAEHaHsBHTLa9hU5Lyx+TM27z1ji/4ZTtx+fv1n/u20f/wWOoeO8o4zv8mnvupTZLC0jBUDoCHua04ODMD1lex6H8CBXBORUpQpjntcX1g9pXzPvZz8K7/cEZMfBnMKkbsRHgDuFJXTKEeAPogFFVALGEURAkrbCKBAACpUSwO7Cmck6IfVuw/46ejD4ckP6egtbzq7+vGu+6XhE7ztVZ8heb8nghB0H80/kBeFHHgAL3L55S/7cgCCMQQj5JVj/cgh8rJCVOlJJhP1/JG/9f8o//pfIT/6NYzf/i2dLOsOJSuOYbLbReQukHtE9R7gNoGToCs0378mZfegHlUH6tAQ2kYgvjL+WqlyXjGPKPKR4P1Hymr0"
    s = s & "0Wm5/bgTzrJ6bLO874FJudwNun2OkKF26jFVQFRVzf633YE38MKLud4HcCBXI1FxjPckjZSOZGQiyvf+U6i20P+iZCZfErgDeD3GfJYa+3ZEPgPhlYKeBB0wZ/wFxIBkIBnS/LRpuzSvFgERckEPqa/u8tPxm6vp7h/25fiPEfTtBvM6UT1myqpjOiuy9tpXhJ/4i39GEYSYVDxYdF5EcvBlvAilHe/3R2MUwXpPVnnKbs60UxCMES/wRz/+r7XfOcbaPV8vAsaoX5YQ7kDkVZjsIWz2Rox5paiuSAgqGoCkjHuk8QQEdag6NESvgNDkAwA0hICvvDpXOe+rkfPVRsA/ota+U7v939Te4GOaZRfUyM76Q/f44tx5QUAFlXDpkODAE3jh5CAH8CIXby1BoutfFRko0neKzTM+9d3fo/qZ/wRPQH3VATmJMXdg8/sx8mrE3I+YV4CsAqgIYEBVhQUd1EAyAFJXAKNnwOzP2VtEAEQFoRAjhZWsa9SICi6UZUedO6J58b7q+MlH5c43jIT3qes6ZWsbuzUSfABUMQdO6PWUg6v/IhfjA5nzOGtRiXoXvMOVE/RHFE+ILxSzisj9IJ+CMZ+ByT4ZsQ8gcgSi8qsYVTEeEZ13AOqyn5eYB0h/s/CSttT2wQjGWjKb5VmWHzWY1xjvP1uq6jNx1b1mtHtk8Gu/1O38wePm43/yk7TsZaAqGJED5b/+chACvAilHQKMBz3ysqK/O6bwKt47Pus//qDqN30v5WhNjEgHMYcw5h6QtyDyJjH2PsTcgshSa7cxk6+K0SAS6pU+PdW4/smgENA6IRj2hACoBrx3eO/RlBlUVbx3BO92A7wvWPMLPs/fFax5OHh35ruqh3eKH/oXfNU/+WkjeQ4hqKqqsZbj99zFx379nYgIa4890XzO2374B6731/GyloMQ4EUu3fEEUfBZJr"
    s = s & "kI48kI/abvIZQ7GBEDnATuA3mNGPtGRF6JmBMg3X12Z5g59C1JoUFE/QmEvSv+JaR2GGIUIYDpiYabxfs3iq8yQTNVHcnZrd3v+MJ/rI+86U7Nl5aoxmNBVRTUFrmiejUfeyDXQA58sBe5SFx5pbLCz+Tn9dgd9yibFZWfCsauYuydiHkjIm9C5AGMuQWRFcDEYn6jU3Xib//vXOpKgNGL5gj3vmn2awINCUKMVGSVoK+UEN5kfHi1DeGWrz1y9+pHu9OCJ8/IE3/6C/B5XH9CWfHwr/4GD/3Df8BDf/+br/clv6HkIAR4EUld8/eZRUXo7Y4pDh9nok4+/dC/UHb+ITqZ4FeHPXw4hpg7MfYNiHySiLkXMacQOURcy+O6LAvaHEMAJASFkJ5pF/mD1JiAS4cAig+O4D0hKFo/NBA0oBp8CH6qGs6o6u8q4Z0efa838pFH7jz+zKf/19/23/G+X5Uv/o7vlHJjU9fPn2e8saV4T3XsCEVZ4jKLtxaAz/++773eX8/LUg48gBehSFAkBHxmWdqa0h075clvQ71DfvY3ADmMyIOIfDIibxZjH0DMLSADlIAmFL9cdCmPmiykTL4K+PQIXKn/H1d7gxiJDwtYiQ9jrDG2ayQ7asS8WpA/ZJBPN2JecXTHd3niGb5JRKebW7r97DlcWWI7hWS9LhDP3/pwRcdxIM9dDnIA11ke/q7vbn7ffPoMxaDH07/1eywdPcIOyv13vRHG61TWGNCcL337IVG9C7GvQeR1IuY+kFOILCW9rd3+fZRfYwFAVZB6OVeQOusfmk17RNI/7RJ+AggJAqJN75AgKdtgjEgYeJXTiPRAnBH7zOqOu/DUV/zJJ0Xy0YkzH60eXL2ZXzp7TowIwYhMK6fvPrLEKzdH2LK63l/Ry1oODMCLSGyeYWyG39rBHj0qQQNUpeI9hJBJ8KcQcy/G"
    s = s & "vFqMfT0i9yPmZJPwk+bf/ZW/Xt2Noppq+bXUGIHaIOjic/P7mj0/AwgIYCQiDGr0sCIWkZ6qOYLwCiP2jZlTle3d92Sm++EN56pzF87xFAO9nU2ZYvDec9/aNmi0Zgfy/MmBAXgRyXd+x7fzH/7nb3LhySdZH+3SzzrCA39ceew3kAsPLyFyJ8gbwTyEmFdi5BaQLoqieATDRcI6SYotdc0/LebSVmSpV/iW692kEGVhmy68oN5L2o8EEE17EkGkIyInBHkNIeRo5Uoma09P37fzuuGfCq/2H+TDq7frcpjIUiwlHBQEXgA5SAJeB2m7/W0pTIbTII/KNm979EH94af/nXzh4VcW3aw3lLx7Oyb7ZBXz6WLMKxFzc8r2c5GE39waHt3/kF5Wu/yK1M0+taiCDxA8GhyqnnmDoBAUDTHZF0LEAeisdTj+HRQfPCEEgkYLpSJTkE2ExwL+XZWOf3nMxgem7DxbFn73prVh+Ib/8jP8tb/79yRHePVrX6c//hM/Tp7nfPzRjzeH8HW//MvX+yt82chBEvBFJM57UQ3SGw5h58O8+djddPP+khh7B/AgxrxarL0PMacRGTZvlKbJZqb88b+W5gqqBtSkWJ1rsMa2scF1+S89I0RPwAhiDcZajEgBeki9v8X76kGv1VtEzesL+qf6bjmrCPygCIPBQPqDARih1+vR7/ev91fzspUDD+AFlH1X/nr1FWHaMWI90pt6nlqt9M0fpIdyB/BajH09NnsDxj4IcugyHzVTflms/UdPQIitvvHnfh5ASG3BHtVULqxX+VTuq72AOuVYewL1dl//nQyFBsV7Twhh6iifcUw/EsT9rqL/I9fid04e/ZS1C7fkbGZjzqvXJ37/vbz3935PL5w/z/c+/TT/4U99CUtLSzzwwAPAQdPQtZADD+A6S0LlCyLygTef4P6v/Wvh5JqTt7"
    s = s & "xnVEjQQ9jsNrLitdj8DamxZ3BlO97jFVzFQZFSAgbEIsamcl8s+WGk+duIwTQ/LaZ+nZgI9xfBxAiEGS+IFqJy1IbsHuuLB3LfudtqdvP6hd85VHzk97LhZ3+KfsqdO/rMs8/K4dVVc/Pp0xL+839laWmJ5eXl6/2VvazkwABcZ1EjoiKCqpybnge+DTvRXLPuKsbegtj7MPY1GPNAzPhfUeK2XvUNz9UAGEAMIhaRLCq81Am+thGwc8ZATAwFovIbDAlmUMMLIo0BomZoNDuVaX6n1fxeg7lbNdxUjsf93/xH70C/6O8RnBObZSwvL/OTP/WTvOf33sMP/fAPXe+v7GUlB1WA51najT2j0QhV5dSpm9ja2kZVRYyB3Cq7U5HfeDfv/M8r2cSeWunYzt1GzEOIeR1i7gBWW7ttw3ufP5Hmn/mfTYuwxiqCyJyrX1cURFNrAS3IgST4kBgUOqocN6L3CrKp4Ceq41FZjuR2qb773j9MN8s4deoU995zr7z1rW/Vsix5/PHHyfN87toeIAWfmxwYgBdQjDGoQhnBLRGHZwTuvEnz//kR7t4UuyJFL1i5ScW+AcxbEe4Hjjc7mY/tn1/Zp/y/n6hqgv/Oqgv7708QBSMGNRCCIKJDg94rgvHgPJxdMnbtM1/39p1l7eiJ0yf1jtO3ijGGCxcuYK1RYyzmoJX4msjBVXwB5cEv/AJe9Re/irNnzxF8hLiMM7jr835c7dTLq+TU4IQMTlm4W+AhRF4PchcwZAbVu4SGPQ9yySxCTRWmTYJwkf9TF3eVQoM6pBCRDiI3gdyP8oBVvfsVmNPfmq8sP5D3zBt/9Ge08+oHVETwzrG9vcNr/+7f4YE3vfEFuwQvZ7HX+wBe7vKlr39D8/vv/fL/4EM//wuIiAxtB5NZeeVfXw1/79FlPn68zG4qi1u6ZK/LxL5JxL5WRG5H"

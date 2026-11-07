Attribute VB_Name = "modVisual"
Option Explicit
' =========================================================
' XLBOOSTER - FULL FIXED VISUAL MODULE
' Focus Cell + QR + Barcode + Utilities Fixed
' =========================================================
Private ribbonUI As IRibbonUI
Public focusColor As Long
Public focusIsOn As Boolean
Public focusRow As Long
Public focusCol As Long
Private prevRow As Range
Private prevCol As Range

Private prevRowColors() As Variant
Private prevColColors() As Variant
' =========================================================
' Ribbon
' =========================================================
Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    Set ribbonUI = ribbon
End Sub
' =========================================================
' HIGHLIGHTS
' =========================================================
Private Sub ApplyHighlightBorder(borderRGB As Long)
    With Selection.Borders
        .LineStyle = xlContinuous
        .color = borderRGB
        .Weight = xlMedium
    End With
End Sub
Public Sub Highlights_Red(control As IRibbonControl)
    ApplyHighlightBorder RGB(255, 0, 0)
End Sub
Public Sub Highlights_Green(control As IRibbonControl)
    ApplyHighlightBorder RGB(0, 176, 80)
End Sub
Public Sub Highlights_Blue(control As IRibbonControl)
    ApplyHighlightBorder RGB(0, 112, 192)
End Sub
Public Sub Highlights_Orange(control As IRibbonControl)
    ApplyHighlightBorder RGB(255, 140, 0)
End Sub
Public Sub Highlights_Clear(control As IRibbonControl)
    Selection.Interior.ColorIndex = xlColorIndexNone
    Selection.Borders.LineStyle = xlNone
End Sub
' =========================================================
' INCREASE / DECREASE COLOUR
' =========================================================

Public Sub Highlights_IncreaseColour(control As IRibbonControl)
    IncreaseColourHelper
End Sub

Public Sub Highlights_DecreaseColour(control As IRibbonControl)
    DecreaseColourHelper
End Sub

Public Sub IncreaseColourHelper()
    Dim c As Range
    Dim r As Long, g As Long, b As Long
    Dim clr As Long
    For Each c In Selection.Cells
        If c.Interior.ColorIndex <> xlNone Then
            clr = c.Interior.color
            b = (clr \ 65536) And 255
            g = (clr \ 256) And 255
            r = clr And 255
            r = Application.WorksheetFunction.Min(r + 30, 255)
            g = Application.WorksheetFunction.Min(g + 30, 255)
            b = Application.WorksheetFunction.Min(b + 30, 255)
            c.Interior.color = RGB(r, g, b)
        End If
    Next c
End Sub

Public Sub DecreaseColourHelper()
    Dim c As Range
    Dim r As Long, g As Long, b As Long
    Dim clr As Long
    For Each c In Selection.Cells
        If c.Interior.ColorIndex <> xlNone Then
            clr = c.Interior.color
            b = (clr \ 65536) And 255
            g = (clr \ 256) And 255
            r = clr And 255
            r = Application.WorksheetFunction.Max(r - 30, 0)
            g = Application.WorksheetFunction.Max(g - 30, 0)
            b = Application.WorksheetFunction.Max(b - 30, 0)
            c.Interior.color = RGB(r, g, b)
        End If
    Next c
End Sub


' =========================================================
' STATUS TICKMARKS
' =========================================================
Private Sub ApplyTickmark(sym As String, bgRGB As Long, fgRGB As Long)
    Dim c As Range
    For Each c In Selection.Cells
        With c
            .Value = sym
            .Interior.color = bgRGB
            .Font.color = fgRGB
            .Font.Name = "Segoe UI"
            .Font.Size = 11
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next c
End Sub

Public Sub Tickmark_Completed(control As IRibbonControl)
    ApplyTickmark ChrW(10004), RGB(198, 239, 206), RGB(0, 97, 0)
End Sub

Public Sub Tickmark_Rejected(control As IRibbonControl)
    ApplyTickmark ChrW(10006), RGB(255, 199, 199), RGB(156, 0, 6)
End Sub

Public Sub Tickmark_NeedsReview(control As IRibbonControl)
    ApplyTickmark "?", RGB(255, 255, 153), RGB(156, 87, 0)
End Sub

Public Sub Tickmark_ApprovedFlag(control As IRibbonControl)
    ApplyTickmark ChrW(10004), RGB(0, 176, 80), RGB(255, 255, 255)
End Sub

Public Sub Tickmark_AttentionRequired(control As IRibbonControl)
    ApplyTickmark ChrW(9888), RGB(255, 192, 0), RGB(156, 87, 0)
End Sub

Public Sub Tickmark_HighPriority(control As IRibbonControl)
    ApplyTickmark "!!", RGB(255, 0, 0), RGB(255, 255, 255)
End Sub

Public Sub Tickmark_Important(control As IRibbonControl)
    ApplyTickmark ChrW(9733), RGB(255, 182, 193), RGB(156, 0, 6)
End Sub

Public Sub Tickmark_Warning(control As IRibbonControl)
    ApplyTickmark ChrW(9888), RGB(255, 235, 156), RGB(156, 87, 0)
End Sub

Public Sub Tickmark_Pending(control As IRibbonControl)
    ApplyTickmark "...", RGB(240, 240, 240), RGB(60, 60, 60)
End Sub

Public Sub Tickmark_InProgress(control As IRibbonControl)
    ApplyTickmark ChrW(9654), RGB(189, 215, 238), RGB(0, 70, 127)
End Sub

Public Sub Tickmark_Information(control As IRibbonControl)
    ApplyTickmark "i", RGB(221, 235, 247), RGB(0, 70, 127)
End Sub

Public Sub Tickmark_KeyRecord(control As IRibbonControl)
    ApplyTickmark ChrW(9733), RGB(255, 230, 153), RGB(156, 101, 0)
End Sub

Public Sub Tickmark_Locked(control As IRibbonControl)
    ApplyTickmark ChrW(9632), RGB(217, 217, 217), RGB(60, 60, 60)
End Sub

Public Sub Tickmark_Blocked(control As IRibbonControl)
    ApplyTickmark "X", RGB(255, 0, 0), RGB(255, 255, 255)
End Sub

Public Sub Tickmark_RemoveAll(control As IRibbonControl)
    Dim c As Range
    For Each c In Selection.Cells
        c.ClearContents
        c.Interior.ColorIndex = xlColorIndexNone
        c.Font.color = RGB(0, 0, 0)
    Next c
End Sub


' =========================================================
' IMAGE IN COMMENT
' =========================================================

Public Sub DataUtilities_ImageInComment(control As IRibbonControl)

    If Selection.Cells.count <> 1 Then
        MsgBox "Select one cell first."
        Exit Sub
    End If

    Dim fPath As String

    fPath = Application.GetOpenFilename( _
    "Images (*.png;*.jpg),*.png;*.jpg")

    If fPath = "False" Then Exit Sub

    Dim scaleInput As String
    scaleInput = InputBox("Enter image scale %", "Image Size", "100")

    If scaleInput = "" Then Exit Sub

    Dim scalePct As Double
    scalePct = val(scaleInput) / 100

    If scalePct <= 0 Then scalePct = 1

    On Error Resume Next
    Selection.Comment.Delete
    On Error GoTo 0

    Dim cmt As Comment

    Set cmt = Selection.AddComment("")

    With cmt.Shape

        .Fill.UserPicture fPath
        .Width = 150 * scalePct
        .Height = 150 * scalePct

    End With

    cmt.Visible = False

    MsgBox "Image added."

End Sub

' =========================================================
' QR CODE FIXED
' =========================================================

Public Sub DataUtilities_CreateQRCode(control As IRibbonControl)

    Dim sizeInput As String

    sizeInput = InputBox( _
        "Enter QR Size", _
        "QR Code", "250")

    If sizeInput = "" Then Exit Sub
    If Not IsNumeric(sizeInput) Then Exit Sub

    Dim qrSize As Long
    qrSize = CLng(sizeInput)

    Dim c As Range

    Application.ScreenUpdating = False

    For Each c In Selection.Cells

        If Trim(c.Value) <> "" Then

            Dim apiURL As String

            apiURL = "https://api.qrserver.com/v1/create-qr-code/?size=" & _
                     qrSize & "x" & qrSize & "&data=" & EncodeURL(CStr(c.Value))

            Dim pic As Picture

            On Error Resume Next
            Set pic = ActiveSheet.Pictures.Insert(apiURL)
            On Error GoTo 0

            If Not pic Is Nothing Then

                With pic
                    .left = c.left
                    .Top = c.Offset(1, 0).Top
                    .Width = 90
                    .Height = 90
                End With

            End If

        End If

    Next c

    Application.ScreenUpdating = True

    MsgBox "QR Code generated."

End Sub

Public Sub DataUtilities_ClearQRCodes(control As IRibbonControl)

    Dim shp As Shape

    For Each shp In ActiveSheet.Shapes

        If shp.Type = msoPicture Then

            If left(shp.Name, 3) = "Pic" Then
                shp.Delete
            End If

        End If

    Next shp

End Sub

Private Function EncodeURL(txt As String) As String

    Dim i As Long
    Dim result As String
    Dim ch As String

    Dim safe As String
    safe = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_.~"

    For i = 1 To Len(txt)

        ch = Mid(txt, i, 1)

        If InStr(safe, ch) > 0 Then
            result = result & ch
        Else
            result = result & "%" & Hex(asc(ch))
        End If

    Next i

    EncodeURL = result

End Function

' =========================================================
' BARCODE - MODERN IMAGE STYLE
' Popup Size + Real Barcode Look
' =========================================================

Public Sub DataUtilities_CreateBarCode(control As IRibbonControl)

    If Selection.Cells.count = 0 Then Exit Sub

    ' -----------------------------------------
    ' Size Popup
    ' -----------------------------------------
    
    Dim sizeInput As String

    sizeInput = InputBox( _
        "Enter desired barcode width in pixels:" & vbLf & _
        "Standard: 250 | High: 350 | Ultra: 500", _
        "Bar Code Size", "350")

    If sizeInput = "" Then Exit Sub

    If Not IsNumeric(sizeInput) Then
        MsgBox "Please enter a valid size.", vbExclamation
        Exit Sub
    End If

    Dim barSize As Long
    barSize = CLng(sizeInput)

    If barSize < 150 Then barSize = 150
    If barSize > 1000 Then barSize = 1000

    Application.ScreenUpdating = False

    Dim c As Range

    For Each c In Selection.Cells

        If Trim(CStr(c.Value)) <> "" Then

            Dim barcodeURL As String

            ' online barcode api
            barcodeURL = "https://barcode.tec-it.com/barcode.ashx?data=" & _
             EncodeURL(CStr(c.Value)) & _
             "&code=Code128&dpi=96&translate-esc=on&hidehrt=True"

            Dim pic As Picture

            On Error Resume Next
            Set pic = ActiveSheet.Pictures.Insert(barcodeURL)
            On Error GoTo 0

            If Not pic Is Nothing Then

                With pic

                    .left = c.Offset(0, 2).left
                    .Top = c.Top

                    ' modern wide barcode look
                    .Width = barSize
                    .Height = barSize * 0.28

                End With

            End If

        End If

    Next c

    Application.ScreenUpdating = True

    MsgBox "Barcode generated successfully.", vbInformation

End Sub

' =========================================================
' TRIM TEXT (ADVANCED)
' Removes:
' 1. Start/End extra spaces
' 2. Multiple spaces inside text
' 3. Hidden non-breaking spaces
' =========================================================

Public Sub DataUtilities_TrimText(control As IRibbonControl)

    Dim c As Range
    Dim txt As String

    If TypeName(Selection) <> "Range" Then Exit Sub

    Application.ScreenUpdating = False

    For Each c In Selection.Cells

        If Not IsEmpty(c.Value) Then

            txt = CStr(c.Value)

            ' remove hidden spaces
            txt = Replace(txt, Chr(160), " ")

            ' clean unwanted characters
            txt = Application.WorksheetFunction.Clean(txt)

            ' proper trim
            txt = Application.WorksheetFunction.Trim(txt)

            c.Value = txt

        End If

    Next c

    Application.ScreenUpdating = True

    MsgBox "Trim complete.", vbInformation

End Sub
' =========================================================
' FILL RANDOM DATA
' Random Numbers / Bangladesh Names
' =========================================================

Public Sub DataUtilities_FillRandom(control As IRibbonControl)

    Dim answer As VbMsgBoxResult

    answer = MsgBox( _
        "What do you want to fill in the blank cells?" & vbLf & vbLf & _
        "Click Yes to fill with random numbers (10 to 1000)" & vbLf & _
        "Click No to fill with random names" & vbLf & _
        "Click Cancel to stop the operation", _
        vbYesNoCancel + vbQuestion, _
        "Fill Random Data")

    If answer = vbCancel Then Exit Sub

    Dim rng As Range

    On Error Resume Next

    Set rng = Application.InputBox( _
        "Please select the range where blank cells should be filled:", _
        "Select Range", Type:=8)

    On Error GoTo 0

    If rng Is Nothing Then Exit Sub

    ' Bangladesh random names
    Dim bdNames As Variant

    bdNames = Split( _
        "Rahim khandakar|Karim miah|Jahid hasan|Sakib khan|Tanvir Rahman|" & _
        "Nadim sorder|Jony miah|Shihab ali hasan|Rakib|Sabbir Hussein|" & _
        "Rafi|Hasan|Naim|Rasel|Mahin|Fahim|Arafat|Tanjim|" & _
        "Md Sharif Khandakar|Alamin|Mim Akter|Nusrat Begum|" & _
        "Jannat Akter|Sumaiya Akter|Mahi Sarkar|Afjal Hassan", "|")

    Randomize

    Dim c As Range

    Application.ScreenUpdating = False

    For Each c In rng.Cells

        If Trim(CStr(c.Value)) = "" Then

            If answer = vbYes Then

                c.Value = Int((1000 - 10 + 1) * Rnd + 10)

            ElseIf answer = vbNo Then

                c.Value = bdNames(Int(Rnd * (UBound(bdNames) + 1)))

            End If

        End If

    Next c

    Application.ScreenUpdating = True

    MsgBox "Blank cells filled successfully.", _
           vbInformation, "Microsoft Excel"

End Sub

' =========================================================
' FILL BLANKS WITH 0 ONLY
' =========================================================

Public Sub DataUtilities_FillBlanks(control As IRibbonControl)

    Dim rng As Range

    On Error Resume Next

    Set rng = Application.InputBox( _
        "Please select the range where blank cells should be filled with 0:", _
        "Fill Blanks", Type:=8)

    On Error GoTo 0

    If rng Is Nothing Then Exit Sub

    Dim c As Range

    Application.ScreenUpdating = False

    For Each c In rng.Cells

        If Trim(CStr(c.Value)) = "" Then

            c.Value = 0

        End If

    Next c

    Application.ScreenUpdating = True

    MsgBox "All blank cells filled with 0 successfully.", _
           vbInformation, "Microsoft Excel"

End Sub

' =========================================================
' INSERT CHECKBOXES - MODERN STYLE
' No TRUE/FALSE
' =========================================================

Public Sub DataUtilities_InsertCheckBoxes(control As IRibbonControl)

    Dim rng As Range

    On Error Resume Next

    Set rng = Application.InputBox( _
        "Select a cell range to insert check boxes", _
        "XLBooster", Type:=8)

    On Error GoTo 0

    If rng Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    Dim c As Range
    Dim cb As CheckBox

    For Each c In rng.Cells

        ' remove old checkbox if exists
        On Error Resume Next
        ActiveSheet.CheckBoxes("chk_" & c.Address(False, False)).Delete
        On Error GoTo 0

        Set cb = ActiveSheet.CheckBoxes.Add( _
            c.left + ((c.Width - 12) / 2), _
            c.Top + ((c.Height - 12) / 2), _
            12, 12)

        With cb

            .Name = "chk_" & c.Address(False, False)

            .Caption = ""

            .LinkedCell = ""

            .Display3DShading = False

            .Value = xlOff

        End With

    Next c

    Application.ScreenUpdating = True

    MsgBox "Check boxes inserted successfully.", vbInformation

End Sub


' =========================================================
' REMOVE CHECKBOXES
' =========================================================

Public Sub DataUtilities_RemoveCheckBoxes(control As IRibbonControl)

    Dim cb As CheckBox

    Application.ScreenUpdating = False

    For Each cb In ActiveSheet.CheckBoxes

        cb.Delete

    Next cb

    Application.ScreenUpdating = True

    MsgBox "All check boxes removed.", vbInformation

End Sub

' =========================================================
' SWAP COLUMNS - FIXED
' =========================================================

Public Sub DataUtilities_SwapColumns(control As IRibbonControl)

    Dim rng1 As Range
    Dim rng2 As Range

    On Error Resume Next

    Set rng1 = Application.InputBox( _
        "Select the first column to swap:", _
        "Input", Type:=8)

    If rng1 Is Nothing Then Exit Sub

    Set rng2 = Application.InputBox( _
        "Select the second column to swap:", _
        "Input", Type:=8)

    If rng2 Is Nothing Then Exit Sub

    On Error GoTo 0

    Dim lastRow As Long

    lastRow = ActiveSheet.UsedRange.Rows.count

    Dim i As Long
    Dim tmp As Variant

    Application.ScreenUpdating = False

    For i = 1 To lastRow

        tmp = Cells(i, rng1.Column).Value

        Cells(i, rng1.Column).Value = Cells(i, rng2.Column).Value

        Cells(i, rng2.Column).Value = tmp

    Next i

    Application.ScreenUpdating = True

    MsgBox "Columns swapped successfully.", vbInformation

End Sub


' =========================================================
' SWAP ROWS - FIXED
' =========================================================

Public Sub DataUtilities_SwapRows(control As IRibbonControl)

    Dim rng1 As Range
    Dim rng2 As Range

    On Error Resume Next

    Set rng1 = Application.InputBox( _
        "Select the first row to swap:", _
        "Input", Type:=8)

    If rng1 Is Nothing Then Exit Sub

    Set rng2 = Application.InputBox( _
        "Select the second row to swap:", _
        "Input", Type:=8)

    If rng2 Is Nothing Then Exit Sub

    On Error GoTo 0

    Dim lastCol As Long

    lastCol = ActiveSheet.UsedRange.Columns.count

    Dim j As Long
    Dim tmp As Variant

    Application.ScreenUpdating = False

    For j = 1 To lastCol

        tmp = Cells(rng1.Row, j).Value

        Cells(rng1.Row, j).Value = Cells(rng2.Row, j).Value

        Cells(rng2.Row, j).Value = tmp

    Next j

    Application.ScreenUpdating = True

    MsgBox "Rows swapped successfully.", vbInformation

End Sub


' =========================================================
' SWAP CELLS - FIXED
' =========================================================

Public Sub DataUtilities_SwapCells(control As IRibbonControl)

    Dim c1 As Range
    Dim c2 As Range

    On Error Resume Next

    Set c1 = Application.InputBox( _
        "Select the first cell to swap:", _
        "Input", Type:=8)

    If c1 Is Nothing Then Exit Sub

    Set c2 = Application.InputBox( _
        "Select the second cell to swap:", _
        "Input", Type:=8)

    If c2 Is Nothing Then Exit Sub

    On Error GoTo 0

    Dim tmp As Variant

    tmp = c1.Value

    c1.Value = c2.Value
    c2.Value = tmp

    MsgBox "Cells swapped successfully.", vbInformation

End Sub


' =========================================================
' FOCUS CELL - SAFE VERSION
' Existing formatting remove hobe na
' =========================================================

Public Sub Visual_FocusCell(control As IRibbonControl)

    If focusIsOn Then

        FocusCell_Clear
        focusIsOn = False

        MsgBox "Focus Cell OFF", vbInformation

    Else

        focusColor = RGB(255, 255, 153)
        focusIsOn = True

        Call ThisWorkbook.ActivateAppEvent

        PaintFocusCell

        MsgBox "Focus Cell ON", vbInformation

    End If

End Sub
' =========================================================
' PAINT FOCUS CELL - PREMIUM VERSION
' =========================================================
Public Sub PaintFocusCell()

    Dim ws As Worksheet
    Dim usedRng As Range
    Dim i As Long
    Dim lastRow As Long
    Dim lastCol As Long

    If Not focusIsOn Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub

    Set ws = ActiveSheet

    On Error Resume Next

    lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    If lastRow = 0 Or lastCol = 0 Then Exit Sub

    Set usedRng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' clear previous focus
    FocusCell_Clear

    ' current row + column
    Set prevRow = Intersect(ws.Rows(ActiveCell.Row), usedRng)
    Set prevCol = Intersect(ws.Columns(ActiveCell.Column), usedRng)

    ' ---------- ROW ----------
    If Not prevRow Is Nothing Then

        ReDim prevRowColors(1 To prevRow.Cells.count)

        For i = 1 To prevRow.Cells.count

            prevRowColors(i) = prevRow.Cells(i).DisplayFormat.Interior.color

            ' header skip
            If prevRow.Cells(i).Row > 1 Then
                prevRow.Cells(i).Interior.color = focusColor
            End If

        Next i

    End If

    ' ---------- COLUMN ----------
    If Not prevCol Is Nothing Then

        ReDim prevColColors(1 To prevCol.Cells.count)

        For i = 1 To prevCol.Cells.count

            prevColColors(i) = prevCol.Cells(i).DisplayFormat.Interior.color

            ' skip header + avoid intersection overwrite
            If prevCol.Cells(i).Row > 1 _
            And prevCol.Cells(i).Row <> ActiveCell.Row Then

                prevCol.Cells(i).Interior.color = focusColor

            End If

        Next i

    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub


' =========================================================
' CLEAR FOCUS CELL - FIX PREVIOUS COLOR STUCK
' =========================================================
' =========================================================
' CLEAR FOCUS CELL
' =========================================================
Public Sub FocusCell_Clear()

    Dim i As Long
    Dim c As Range

    On Error Resume Next

    ' restore row
    If Not prevRow Is Nothing Then

        For i = 1 To prevRow.Cells.count
            prevRow.Cells(i).Interior.color = prevRowColors(i)
        Next i

    End If

    ' restore column
    If Not prevCol Is Nothing Then

        For i = 1 To prevCol.Cells.count

            Set c = prevCol.Cells(i)

            If prevRow Is Nothing Then
                c.Interior.color = prevColColors(i)
            Else
                If Intersect(c, prevRow) Is Nothing Then
                    c.Interior.color = prevColColors(i)
                End If
            End If

        Next i

    End If

    Set prevRow = Nothing
    Set prevCol = Nothing

    On Error GoTo 0

End Sub


' =========================================================
' COLOR CHANGE
' =========================================================
Public Sub Visual_FocusCellColor(control As IRibbonControl)

    If Not focusIsOn Then Exit Sub

    focusColor = RGB(Int(Rnd * 255), _
                     Int(Rnd * 255), _
                     Int(Rnd * 255))

    PaintFocusCell

    MsgBox "Focus color changed.", vbInformation

End Sub


' =========================================================
' PRESET COLORS
' =========================================================
Private Sub SetFocusColor(clr As Long)

    focusColor = clr

    If focusIsOn Then
        PaintFocusCell
    End If

End Sub


Public Sub FocusCell_SetYellow(control As IRibbonControl)
    SetFocusColor RGB(255, 255, 153)
End Sub

Public Sub FocusCell_SetLightBlue(control As IRibbonControl)
    SetFocusColor RGB(173, 216, 230)
End Sub

Public Sub FocusCell_SetLightGreen(control As IRibbonControl)
    SetFocusColor RGB(198, 239, 206)
End Sub

Public Sub FocusCell_SetOrange(control As IRibbonControl)
    SetFocusColor RGB(255, 229, 153)
End Sub

Public Sub FocusCell_SetPink(control As IRibbonControl)
    SetFocusColor RGB(255, 199, 206)
End Sub

Public Sub FocusCell_SetPurple(control As IRibbonControl)
    SetFocusColor RGB(218, 179, 255)
End Sub

Public Sub FocusCell_SetCyan(control As IRibbonControl)
    SetFocusColor RGB(153, 255, 255)
End Sub


' =========================================================
' TURN OFF BUTTON
' =========================================================
Public Sub FocusCell_TurnOffBtn(control As IRibbonControl)

    FocusCell_Clear
    focusIsOn = False

    MsgBox "Focus Cell OFF", vbInformation

End Sub


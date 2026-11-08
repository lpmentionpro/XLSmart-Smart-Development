VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDashboard 
   Caption         =   "Power Dashboard"
   ClientHeight    =   9105.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7425
   OleObjectBlob   =   "frmDashboard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mHeaderRow As Long

Private Sub UserForm_Initialize()

    Me.Caption = "Power Dashboard Builder"

    With cmbThemeComboBox
        .Clear
        .AddItem "Ocean (Blue Light)"
        .AddItem "Midnight (Dark)"
        .AddItem "Slate (Dark Blue)"
        .AddItem "Forest (Green)"
        .AddItem "Rose (Pink)"
        .AddItem "Navy (Deep Blue)"
        .ListIndex = 1
    End With

    With cmbCurrencyComboBox
        .Clear
        .AddItem "None"
        .AddItem "Bangladesh (BDT)"
        .AddItem "India (INR)"
        .AddItem "Europe (EUR)"
        .AddItem "International (USD)"
        .ListIndex = 1
    End With

    With cmbGridlinesComboBox
        .Clear
        .AddItem "Show"
        .AddItem "Hide"
        .ListIndex = 0
    End With

    With cmbDashTypeComboBox
        .Clear
        .AddItem "Web Dashboard (HTML)"
        .ListIndex = 0
    End With

    txtTitleTextBox.text = "My Excel Dashboard"

    radiooptSingleOptionButton.Value = True
    chkSmartLockCheckBox.Value = True

    txtSheets.text = ""
    txtSheets.Visible = False

    txtRangeTextBox.text = ""
    txtRangeTextBox.Tag = ""

    mHeaderRow = 1

    UpdateDataSourceUI

    ' Auto-detect data range from ActiveCell (Ctrl+Shift+Down+Right)
    On Error Resume Next
    If Not ActiveCell Is Nothing Then
        If ActiveCell.CurrentRegion.Rows.count >= 2 And ActiveCell.CurrentRegion.Columns.count >= 1 Then
            Dim initRng As Range
            Set initRng = ActiveCell.CurrentRegion
            txtRangeTextBox.Value = initRng.Worksheet.Name & "!" & initRng.Address(False, False)
            txtRangeTextBox.Tag = initRng.Worksheet.Parent.Name & "|" & initRng.Worksheet.Name & "|" & initRng.Address(False, False)
            lblStatusLabel.Caption = "Auto-detected: " & (initRng.Rows.count - 1) & " data rows"
        End If
    End If
    On Error GoTo 0

End Sub

Private Sub radiooptSingleOptionButton_Click()
    UpdateDataSourceUI
End Sub

Private Sub radiooptMultiOptionButton_Click()
    UpdateDataSourceUI
End Sub

Private Sub UpdateDataSourceUI()

    txtRangeTextBox.Visible = True
    btnPickRangeCommandButton.Visible = True
    txtSheets.Visible = False

    txtRangeTextBox.text = ""
    txtRangeTextBox.Tag = ""

    txtSheets.text = ""
    mHeaderRow = 1

    If radiooptSingleOptionButton.Value = True Then
        lblStatusLabel.Caption = _
            "Click Pick and select any cell inside your data table."
    Else
        lblStatusLabel.Caption = _
            "Click Pick and select sheets."
    End If

End Sub

Private Sub btnPickRangeCommandButton_Click()

    If radiooptSingleOptionButton.Value Then
        PickSingleTableRange
    Else
        PickMultipleSheets
    End If

End Sub

Private Sub PickSingleTableRange()

    Dim selected_rng As Range
    Dim final_rng As Range
    Dim defaultAddr As String

    ' Smart default from ActiveCell (Ctrl+Shift+Down+Right region)
    On Error Resume Next
    If Not ActiveCell Is Nothing Then
        If ActiveCell.CurrentRegion.Cells.count > 1 Then
            defaultAddr = ActiveCell.CurrentRegion.Address(False, False)
        Else
            defaultAddr = ActiveCell.Address(False, False)
        End If
    End If
    On Error GoTo 0

    Me.Hide
    DoEvents

    On Error Resume Next
    Set selected_rng = Application.InputBox( _
        prompt:="Click any cell inside your table or select range." & vbCrLf & _
                "Tip: Click 1 cell or header to auto-select whole table (Ctrl+Shift+Down+Right).", _
        Title:="Select Data Range", _
        Default:=defaultAddr, _
        Type:=8)
    On Error GoTo 0

    If selected_rng Is Nothing Then
        lblStatusLabel.Caption = "No range selected."
        Me.Show
        Exit Sub
    End If

    If selected_rng.Areas.count > 1 Then
        MsgBox "Please select one continuous range only.", vbExclamation
        Me.Show
        Exit Sub
    End If

    ' Smart auto-expand (Ctrl+Shift+Down+Right behavior)
    Dim ws As Worksheet
    Set ws = selected_rng.Worksheet
    Dim startCell As Range
    Set startCell = selected_rng.Cells(1, 1)

    If selected_rng.Cells.count = 1 Then
        ' 1 cell clicked -> auto-expand to full CurrentRegion
        If startCell.CurrentRegion.Rows.count >= 2 Then
            Set final_rng = startCell.CurrentRegion
        Else
            ' Fallback to Ctrl+Shift+Down & Ctrl+Shift+Right
            Dim lastR As Long, lastC As Long
            lastR = ws.Cells(ws.Rows.count, startCell.Column).End(xlUp).Row
            lastC = ws.Cells(startCell.Row, ws.Columns.count).End(xlToLeft).Column
            If lastR >= startCell.Row And lastC >= startCell.Column Then
                Set final_rng = ws.Range(startCell, ws.Cells(lastR, lastC))
            Else
                Set final_rng = startCell
            End If
        End If
    ElseIf selected_rng.Rows.count = 1 And selected_rng.Columns.count > 1 Then
        ' Header row selected -> auto-expand down (Ctrl+Shift+Down)
        Dim maxR As Long
        maxR = ws.Cells(ws.Rows.count, startCell.Column).End(xlUp).Row
        If maxR > startCell.Row Then
            Set final_rng = ws.Range(startCell, ws.Cells(maxR, selected_rng.Cells(1, selected_rng.Columns.count).Column))
        Else
            Set final_rng = selected_rng
        End If
    Else
        Set final_rng = selected_rng
    End If

    If final_rng.Rows.count < 2 Then
        MsgBox "Please include headers and at least one data row.", vbExclamation
        Me.Show
        Exit Sub
    End If

    txtRangeTextBox.Value = _
        final_rng.Worksheet.Name & "!" & _
        final_rng.Address(False, False)

    ' Workbook|Sheet|Range
    txtRangeTextBox.Tag = _
        final_rng.Worksheet.Parent.Name & "|" & _
        final_rng.Worksheet.Name & "|" & _
        final_rng.Address(False, False)

    lblStatusLabel.Caption = _
        "Range Selected: " & _
        final_rng.Rows.count - 1 & _
        " data rows"

    Me.Show

End Sub

Private Sub PickMultipleSheets()

    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation
        Exit Sub
    End If

    frmSheetPicker.Show vbModal

    If frmSheetPicker.WasCancelled Then

        lblStatusLabel.Caption = "No sheets selected."
        Unload frmSheetPicker
        Exit Sub

    End If

    txtSheets.text = frmSheetPicker.SelectedSheetIndexes
    mHeaderRow = frmSheetPicker.headerRow

    txtRangeTextBox.Tag = txtSheets.text

    txtRangeTextBox.text = _
        "Multi Sheet (" & _
        CountSelectedSheets(txtSheets.text) & _
        " selected)"

    lblStatusLabel.Caption = _
        CountSelectedSheets(txtSheets.text) & _
        " sheets selected"

    Unload frmSheetPicker

End Sub

Private Function CountSelectedSheets(ByVal sheetText As String) As Long

    If Trim(sheetText) = "" Then
        CountSelectedSheets = 0
    Else
        CountSelectedSheets = _
            UBound(Split(sheetText, ",")) + 1
    End If

End Function

Private Sub chkSmartLockCheckBox_Click()

    If chkSmartLockCheckBox.Value Then
        lblStatusLabel.Caption = _
            "Smart Lock enabled."
    Else
        lblStatusLabel.Caption = _
            "Smart Lock disabled."
    End If

End Sub

Private Sub btnCreateCommandButton_Click()

    Dim themeName As String
    Dim currencyName As String
    Dim dataInput As String

    If Trim(txtTitleTextBox.text) = "" Then
        MsgBox "Please enter dashboard title.", vbExclamation
        txtTitleTextBox.SetFocus
        Exit Sub
    End If

    If radiooptSingleOptionButton.Value Then

        If Trim(txtRangeTextBox.Tag) = "" Then
            MsgBox "Please select data range first.", vbExclamation
            Exit Sub
        End If

        dataInput = txtRangeTextBox.Tag
        mHeaderRow = 1

    Else

        If Trim(txtSheets.text) = "" Then
            MsgBox "Please select sheets first.", vbExclamation
            Exit Sub
        End If

        dataInput = txtSheets.text

    End If

    Select Case cmbThemeComboBox.ListIndex
        Case 0: themeName = "ocean"
        Case 1: themeName = "midnight"
        Case 2: themeName = "slate"
        Case 3: themeName = "forest"
        Case 4: themeName = "rose"
        Case 5: themeName = "navy"
        Case Else: themeName = "midnight"
    End Select

    Select Case cmbCurrencyComboBox.ListIndex
        Case 0: currencyName = ""
        Case 1: currencyName = "BDT"
        Case 2: currencyName = "INR"
        Case 3: currencyName = "EUR"
        Case 4: currencyName = "USD"
        Case Else: currencyName = ""
    End Select

    lblStatusLabel.Caption = "Creating dashboard..."
    btnCreateCommandButton.Enabled = False

    Me.Hide

    On Error GoTo ErrHandler

    CreateDashboard _
    Trim(txtTitleTextBox.text), _
    themeName, _
    currencyName, _
    radiooptMultiOptionButton.Value, _
    dataInput, _
    mHeaderRow, _
    chkSmartLockCheckBox.Value, _
    (cmbGridlinesComboBox.ListIndex = 0)

    Unload Me
    Exit Sub

ErrHandler:

    MsgBox "Dashboard create failed: " & Err.Description, vbCritical

    btnCreateCommandButton.Enabled = True
    lblStatusLabel.Caption = "Error creating dashboard."

    Me.Show

End Sub

Private Sub btnCancelCommandButton_Click()
    Unload Me
End Sub


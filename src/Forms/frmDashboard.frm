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
        .AddItem "Animated (Live Portfolio)"
        .AddItem "Ocean (Blue Light)"
        .AddItem "Midnight (Dark)"
        .AddItem "Slate (Dark Blue)"
        .AddItem "Forest (Green)"
        .AddItem "Rose (Pink)"
        .AddItem "Navy (Deep Blue)"
        .ListIndex = 0
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

    btnPickRangeCommandButton.Caption = "Pick Range"

    txtSheets.text = ""
    mHeaderRow = 1

    If radiooptSingleOptionButton.Value = True Then
        lblStatusLabel.Caption = _
            "Click Pick Range, then use mouse or Ctrl+Shift+Arrow in the Excel range box."
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

    ' Type 8 opens Excel's native reference picker. It keeps the worksheet active,
    ' so mouse selection and Ctrl+Shift+Arrow work before the user confirms with Enter.
    Me.Hide
    On Error Resume Next
    Set selected_rng = Application.InputBox( _
        Prompt:="Select the full dashboard table, including the header row. " & _
                "Use mouse or Ctrl+Shift+Arrow, then press Enter.", _
        Title:="Pick Dashboard Data Range", _
        Type:=8)
    Err.Clear
    On Error GoTo 0
    Me.Show vbModeless

    If selected_rng Is Nothing Then
        lblStatusLabel.Caption = "Range selection cancelled."
        Exit Sub
    End If

    If selected_rng.Areas.Count > 1 Then
        MsgBox "Please select one continuous range only.", vbExclamation
        Exit Sub
    End If

    If selected_rng.Rows.Count < 2 Or selected_rng.Columns.Count < 1 Then
        MsgBox "Select the header row and at least one data row.", vbExclamation
        Exit Sub
    End If

    txtRangeTextBox.Value = selected_rng.Worksheet.Name & "!" & _
                            selected_rng.Address(False, False)
    txtRangeTextBox.Tag = selected_rng.Worksheet.Parent.Name & "|" & _
                          selected_rng.Worksheet.Name & "|" & _
                          selected_rng.Address(False, False)

    lblStatusLabel.Caption = "Range Selected: " & _
                             (selected_rng.Rows.Count - 1) & _
                             " data rows (" & selected_rng.Address(False, False) & ")"

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
        Case 0: themeName = "animated"
        Case 1: themeName = "ocean"
        Case 2: themeName = "midnight"
        Case 3: themeName = "slate"
        Case 4: themeName = "forest"
        Case 5: themeName = "rose"
        Case 6: themeName = "navy"
        Case Else: themeName = "animated"
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


Attribute VB_Name = "modValidationTools"
Option Explicit

' ============================================================
' MODULE: modValidationTools
' XLSmart - Control Group - Validation Tools
' ============================================================

Public Sub ValidationTools_Main(control As IRibbonControl)
    MsgBox "Validation Tools - Select a specific option from the dropdown menu.", _
           vbInformation, "XLSmart - Validation Tools"
End Sub

' ----- Block Duplicates -----
Public Sub ValidationTools_BlockDuplicates(control As IRibbonControl)
    Dim rng As Range
    Set rng = GetSelectedRange()
    If rng Is Nothing Then Exit Sub
    
    On Error GoTo ErrHandler
    With rng.Validation
        .Delete
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=COUNTIF(" & rng.Columns(1).Address & "," & rng.Columns(1).Cells(1).Address(False, False) & ")=1"
        .IgnoreBlank = True
        .ShowInput = False
        .ShowError = True
        .ErrorTitle = "Duplicate Entry"
        .ErrorMessage = "Duplicate entries are not allowed in this column!"
    End With
    MsgBox "Block Duplicates validation applied successfully!", vbInformation, "XLSmart"
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "XLSmart"
End Sub

' ----- Gmail Validation -----
Public Sub ValidationTools_Email(control As IRibbonControl)

    Dim rng As Range
    Dim firstCell As String

    Set rng = GetSelectedRange()
    If rng Is Nothing Then Exit Sub

    firstCell = rng.Cells(1, 1).Address(False, False)

    On Error GoTo ErrHandler

    rng.Validation.Delete

    rng.Validation.Add _
        Type:=xlValidateCustom, _
        AlertStyle:=xlValidAlertStop, _
        Formula1:="=RIGHT(LOWER(" & firstCell & "),10)=""@gmail.com"""

    With rng.Validation
        .IgnoreBlank = True
        .ShowError = True
        .ErrorTitle = "Invalid Gmail"
        .ErrorMessage = "Only Gmail addresses are allowed"
    End With

    MsgBox "Email validation added successfully.", vbInformation, "XLSmart"
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "XLSmart"

End Sub

' ----- Bangladesh Phone Validation -----
Public Sub ValidationTools_Phone(control As IRibbonControl)

    Dim rng As Range
    Dim firstCell As String

    Set rng = GetSelectedRange()
    If rng Is Nothing Then Exit Sub

    firstCell = rng.Cells(1, 1).Address(False, False)

    On Error GoTo ErrHandler

    ' keep leading zero
    rng.NumberFormat = "@"

    rng.Validation.Delete

    rng.Validation.Add _
        Type:=xlValidateCustom, _
        AlertStyle:=xlValidAlertStop, _
        Formula1:="=AND(LEN(" & firstCell & "&"""")=11,LEFT(" & firstCell & ",2)=""01"")"

    With rng.Validation
        .IgnoreBlank = True
        .ShowError = True
        .ErrorTitle = "Invalid Phone Number"
        .ErrorMessage = "Enter a valid Bangladesh mobile number (01XXXXXXXXX)"
    End With

    MsgBox "Bangladesh phone validation applied successfully!", vbInformation, "XLSmart"
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "XLSmart"

End Sub
' ----- Numeric Only -----
Public Sub ValidationTools_NumericOnly(control As IRibbonControl)
    Dim rng As Range
    Set rng = GetSelectedRange()
    If rng Is Nothing Then Exit Sub
    
    On Error GoTo ErrHandler
    With rng.Validation
        .Delete
        .Add Type:=xlValidateDecimal, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="-9999999999", _
             Formula2:="9999999999"
        .IgnoreBlank = True
        .ShowError = True
        .ErrorTitle = "Invalid Input"
        .ErrorMessage = "Only numeric values are allowed in this cell!"
    End With
    MsgBox "Numeric Only validation applied successfully!", vbInformation, "XLSmart"
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "XLSmart"
End Sub

' ----- Text Only -----
Public Sub ValidationTools_TextOnly(control As IRibbonControl)

    Dim rng As Range
    Dim firstCell As String

    Set rng = GetSelectedRange()
    If rng Is Nothing Then Exit Sub

    firstCell = rng.Cells(1, 1).Address(False, False)

    On Error GoTo ErrHandler

    rng.Validation.Delete

    rng.Validation.Add _
        Type:=xlValidateCustom, _
        AlertStyle:=xlValidAlertStop, _
        Formula1:="=NOT(ISNUMBER(" & firstCell & "))"

    With rng.Validation
        .IgnoreBlank = True
        .ShowError = True
        .ErrorTitle = "Invalid Input"
        .ErrorMessage = "Only text values are allowed."
    End With

    MsgBox "Text validation applied successfully!", vbInformation, "XLSmart"
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "XLSmart"

End Sub

' ----- URL Validation -----
Public Sub ValidationTools_URL(control As IRibbonControl)

    Dim rng As Range
    Dim firstCell As String

    Set rng = GetSelectedRange()
    If rng Is Nothing Then Exit Sub

    firstCell = rng.Cells(1, 1).Address(False, False)

    On Error GoTo ErrHandler

    rng.Validation.Delete

    rng.Validation.Add _
        Type:=xlValidateCustom, _
        AlertStyle:=xlValidAlertStop, _
        Formula1:="=OR(LEFT(LOWER(" & firstCell & "),7)=""http://"",LEFT(LOWER(" & firstCell & "),8)=""https://"")"

    With rng.Validation
        .IgnoreBlank = True
        .ShowError = True
        .ErrorTitle = "Invalid URL"
        .ErrorMessage = "Enter a valid URL starting with http:// or https://"
    End With

    MsgBox "URL validation applied successfully!", vbInformation, "XLSmart"
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "XLSmart"

End Sub

' ----- Date Validation -----
Public Sub ValidationTools_Date(control As IRibbonControl)
    Dim rng As Range
    Set rng = GetSelectedRange()
    If rng Is Nothing Then Exit Sub
    
    On Error GoTo ErrHandler
    With rng.Validation
        .Delete
        .Add Type:=xlValidateDate, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="1/1/1900", _
             Formula2:="12/31/2099"
        .IgnoreBlank = True
        .ShowError = True
        .ErrorTitle = "Invalid Date"
        .ErrorMessage = "Please enter a valid date between 1900 and 2099."
    End With
    MsgBox "Date validation applied successfully!", vbInformation, "XLSmart"
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "XLSmart"
End Sub

' ----- PIN Validation (6 digit) -----
Public Sub ValidationTools_PIN(control As IRibbonControl)
    Dim rng As Range
    Set rng = GetSelectedRange()
    If rng Is Nothing Then Exit Sub
    
    On Error GoTo ErrHandler
    With rng.Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="100000", _
             Formula2:="999999"
        .IgnoreBlank = True
        .ShowError = True
        .ErrorTitle = "Invalid PIN"
        .ErrorMessage = "Please enter a valid 6-digit PIN code."
    End With
    MsgBox "PIN validation (6 digits) applied successfully!", vbInformation, "XLSmart"
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "XLSmart"
End Sub

' ----- Remove Validation -----
Public Sub ValidationTools_Remove(control As IRibbonControl)
    Dim rng As Range
    Set rng = GetSelectedRange()
    If rng Is Nothing Then Exit Sub
    
    On Error GoTo ErrHandler
    rng.Validation.Delete
    MsgBox "Validation removed successfully", vbInformation, "XLSmart"
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "XLSmart"
End Sub

' ============================================================
' HELPER: Get Selected Range (with guard)
' ============================================================
Private Function GetSelectedRange() As Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell range first!", vbExclamation, "XLSmart"
        Set GetSelectedRange = Nothing
        Exit Function
    End If
    Set GetSelectedRange = Selection
End Function

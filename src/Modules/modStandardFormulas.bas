Attribute VB_Name = "modStandardFormulas"
'==============================================================================
' XLSmart - Standard Formulas Module
' 11 Custom Worksheet Functions (UDFs)
' Compatible with: Excel 2016+
'==============================================================================

Option Explicit

'------------------------------------------------------------------------------
' 1. EXTRACT TEXT ONLY
' Usage: =XL_ExtractText(A1)
' Example: "SKU: 998-X (Wireless Mouse)" -> "SKU X Wireless Mouse"
'------------------------------------------------------------------------------
Function XL_ExtractText(cell As Range) As String
    Dim s As String
    Dim result As String
    Dim i As Integer
    Dim c As String
    
    s = CStr(cell.Value)
    result = ""
    
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Or c = " " Then
            result = result & c
        End If
    Next i
    
    ' Clean up extra spaces
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    XL_ExtractText = Trim(result)
End Function

'------------------------------------------------------------------------------
' 2. EXTRACT NUMBERS ONLY
' Usage: =XL_ExtractNumbers(A1)
' Example: "Invoice Amount: USD 4,500" -> "4500"
'------------------------------------------------------------------------------
Function XL_ExtractNumbers(cell As Range) As String
    Dim s As String
    Dim result As String
    Dim i As Integer
    Dim c As String
    
    s = CStr(cell.Value)
    result = ""
    
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c >= "0" And c <= "9" Then
            result = result & c
        End If
    Next i
    
    If result = "" Then
        XL_ExtractNumbers = ""
    Else
        XL_ExtractNumbers = result
    End If
End Function

'------------------------------------------------------------------------------
' 3. EXTRACT EMAILS (Multiple email support)
' Usage: =XL_ExtractEmails(A1)
' Example: "CC: manager@hr.com, director@hr.com" -> "manager@hr.com, director@hr.com"
'------------------------------------------------------------------------------
Function XL_ExtractEmails(cell As Range) As String
    Dim s As String
    Dim words() As String
    Dim result As String
    Dim i As Integer
    Dim w As String
    
    On Error GoTo ErrHandler
    
    ' cell.Text use koro - value er bodole
    s = cell.text
    If s = "" Then s = CStr(cell.Value)
    
    ' Replace common separators with space
    s = Replace(s, ",", " ")
    s = Replace(s, ";", " ")
    s = Replace(s, "|", " ")
    s = Replace(s, Chr(10), " ")  ' line break handle
    
    words = Split(s, " ")
    result = ""
    
    For i = 0 To UBound(words)
        w = Trim(words(i))
        If Len(w) < 5 Then GoTo NextWord  ' too short to be email
        
        Dim atPos As Integer
        atPos = InStr(w, "@")
        
        If atPos > 1 Then  ' @ must not be first char
            Dim dotPos As Integer
            dotPos = InStr(atPos, w, ".")
            If dotPos > atPos + 1 And dotPos < Len(w) Then  ' dot must be after @ and not last
                ' Clean trailing punctuation
                Do While Len(w) > 0
                    Dim lastChar As String
                    lastChar = right(w, 1)
                    If lastChar = "." Or lastChar = "," Or lastChar = "!" Or _
                       lastChar = ")" Or lastChar = ">" Then
                        w = left(w, Len(w) - 1)
                    Else
                        Exit Do
                    End If
                Loop
                
                If result = "" Then
                    result = w
                Else
                    result = result & ", " & w
                End If
            End If
        End If
        
NextWord:
    Next i
    
    If result = "" Then
        XL_ExtractEmails = "No email found"
    Else
        XL_ExtractEmails = result
    End If
    Exit Function
    
ErrHandler:
    XL_ExtractEmails = "Error: " & Err.Description
End Function
'------------------------------------------------------------------------------
' 4. INDIAN MOBILE NUMBERS (10 digit)
' Usage: =XL_MobileNumberOnly(A1)
' Example: "Call me at +91-9899001122 urgent." -> "9899001122"
' Rules: 10 digit number starting with 6,7,8,9
'------------------------------------------------------------------------------
Function XL_MobileNumberOnly(cell As Range) As String
    Dim s As String
    Dim result As String
    Dim digits As String
    Dim i As Integer
    Dim c As String

    On Error GoTo ErrHandler

    s = cell.text
    If s = "" Then s = CStr(cell.Value)

    ' +880 / +88 strip, 0 rekhe dao
    s = Replace(s, "+880", "0")
    s = Replace(s, "880-", "0")
    s = Replace(s, "+88", "0")
    s = Replace(s, "88-", "0")

    result = ""
    digits = ""

    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c >= "0" And c <= "9" Then
            digits = digits & c
        Else
            result = CheckBD(digits, result)
            digits = ""
        End If
    Next i
    result = CheckBD(digits, result)

    If result = "" Then
        XL_MobileNumberOnly = "No valid number"
    Else
        XL_MobileNumberOnly = result
    End If
    Exit Function

ErrHandler:
    XL_MobileNumberOnly = "Error: " & Err.Description
End Function

Private Function CheckBD(digits As String, existing As String) As String
    Dim matched As String
    matched = ""

    ' BANGLADESH ONLY: 11 digit, 013~019 diye shuru
    If Len(digits) = 11 Then
        Dim prefix3 As String
        prefix3 = left(digits, 3)
        If prefix3 = "013" Or prefix3 = "014" Or prefix3 = "015" Or _
           prefix3 = "016" Or prefix3 = "017" Or prefix3 = "018" Or prefix3 = "019" Then
            matched = digits
        End If
    End If

    If matched = "" Then
        CheckBD = existing
    ElseIf existing = "" Then
        CheckBD = matched
    Else
        CheckBD = existing & ", " & matched
    End If
End Function

'------------------------------------------------------------------------------
' 5. REMOVE SPECIAL CHARACTERS
' Usage: =XL_RemoveSpecialChars(A1)
' Example: "Project_Name_Final@2024!" -> "ProjectNameFinal2024"
' Keeps: A-Z, a-z, 0-9, space
'------------------------------------------------------------------------------
Function XL_RemoveSpecialChars(cell As Range) As String
    Dim s As String
    Dim result As String
    Dim i As Integer
    Dim c As String
    
    s = CStr(cell.Value)
    result = ""
    
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Or _
           (c >= "0" And c <= "9") Or c = " " Then
            result = result & c
        End If
    Next i
    
    ' Clean extra spaces
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    XL_RemoveSpecialChars = Trim(result)
End Function

'------------------------------------------------------------------------------
' 6. EXTRACT URL
' Usage: =XL_ExtractURL(A1)
' Example: "Visit our site https://www.google.com for info." -> "https://www.google.com"
'------------------------------------------------------------------------------
Function XL_ExtractURL(cell As Range) As String
    Dim s As String
    Dim words() As String
    Dim i As Integer
    Dim w As String
    Dim result As String
    Dim lw As String
    
    s = CStr(cell.Value)
    words = Split(s, " ")
    result = ""
    
    For i = 0 To UBound(words)
        w = Trim(words(i))
        lw = LCase(w)
        
        ' Check for http/https prefix or www prefix or common domain patterns
        Dim isURL As Boolean
        isURL = False
        
        If left(lw, 7) = "http://" Or left(lw, 8) = "https://" Then
            isURL = True
        ElseIf left(lw, 4) = "www." Then
            isURL = True
        ElseIf InStr(lw, ".com") > 0 Or InStr(lw, ".net") > 0 Or _
               InStr(lw, ".org") > 0 Or InStr(lw, ".io") > 0 Then
            ' Check it has a slash or dot structure (not just IP)
            If InStr(w, ".") > 0 And Len(w) > 5 Then
                ' Exclude plain IPs like 192.168.1.1
                Dim firstPart As String
                firstPart = left(w, InStr(w, ".") - 1)
                If Not IsNumeric(firstPart) Then
                    isURL = True
                End If
            End If
        End If
        
        If isURL Then
            ' Strip trailing punctuation
            Do While Len(w) > 0 And (right(w, 1) = "." Or right(w, 1) = "," Or right(w, 1) = ")")
                w = left(w, Len(w) - 1)
            Loop
            If result = "" Then
                result = w
            Else
                result = result & ", " & w
            End If
        End If
    Next i
    
    If result = "" Then
        XL_ExtractURL = "No URL found"
    Else
        XL_ExtractURL = result
    End If
End Function

'------------------------------------------------------------------------------
' 7. AGE CALCULATOR FROM DATE OF BIRTH
' Usage: =XL_AgeFromDOB(A1)
' Example: Date serial or date value -> "25 Years 3 Months 12 Days"
'------------------------------------------------------------------------------
Function XL_AgeFromDOB(cell As Range) As String
    Dim dob As Date
    Dim today As Date
    Dim years As Integer
    Dim months As Integer
    Dim days As Integer
    Dim tempDate As Date
    
    On Error GoTo InvalidDate
    
    dob = CDate(cell.Value)
    today = Date
    
    If dob > today Then
        XL_AgeFromDOB = "Future date!"
        Exit Function
    End If
    
    years = DateDiff("yyyy", dob, today)
    If DateSerial(Year(today), Month(dob), Day(dob)) > today Then
        years = years - 1
    End If
    
    tempDate = DateAdd("yyyy", years, dob)
    months = DateDiff("m", tempDate, today)
    If DateAdd("m", months, tempDate) > today Then
        months = months - 1
    End If
    
    tempDate = DateAdd("m", months, DateAdd("yyyy", years, dob))
    days = DateDiff("d", tempDate, today)
    
    XL_AgeFromDOB = years & " Yrs " & months & " Mon " & days & " Days"
    Exit Function
    
InvalidDate:
    XL_AgeFromDOB = "Invalid Date"
End Function

'------------------------------------------------------------------------------
' 8. NUMBER TO WORDS (Currency - Indian style)
' Usage: =XL_ConvertToWords(A1)
' Example: 1500 -> "One Thousand Five Hundred Only"
'------------------------------------------------------------------------------
Function XL_ConvertToWords(cell As Range) As String
    Dim num As Long
    Dim result As String

    On Error GoTo InvalidNum

    If IsEmpty(cell) Or Trim(cell.text) = "" Then
        XL_ConvertToWords = "No Value"
        Exit Function
    End If

    If Not IsNumeric(cell.Value) Then
        XL_ConvertToWords = "Invalid Number"
        Exit Function
    End If

    num = CLng(cell.Value)

    If num = 0 Then
        XL_ConvertToWords = "Zero Only"
        Exit Function
    End If

    result = NumToWords(num)
    XL_ConvertToWords = result & " Only"
    Exit Function

InvalidNum:
    XL_ConvertToWords = "Invalid Number"
End Function

Private Function NumToWords(ByVal num As Long) As String
    Dim ones(19) As String
    Dim tens(9) As String
    Dim result As String

    ones(0) = ""
    ones(1) = "One"
    ones(2) = "Two"
    ones(3) = "Three"
    ones(4) = "Four"
    ones(5) = "Five"
    ones(6) = "Six"
    ones(7) = "Seven"
    ones(8) = "Eight"
    ones(9) = "Nine"
    ones(10) = "Ten"
    ones(11) = "Eleven"
    ones(12) = "Twelve"
    ones(13) = "Thirteen"
    ones(14) = "Fourteen"
    ones(15) = "Fifteen"
    ones(16) = "Sixteen"
    ones(17) = "Seventeen"
    ones(18) = "Eighteen"
    ones(19) = "Nineteen"

    tens(0) = ""
    tens(1) = ""
    tens(2) = "Twenty"
    tens(3) = "Thirty"
    tens(4) = "Forty"
    tens(5) = "Fifty"
    tens(6) = "Sixty"
    tens(7) = "Seventy"
    tens(8) = "Eighty"
    tens(9) = "Ninety"

    result = ""

    If num >= 10000000 Then
        result = result & NumToWords(num \ 10000000) & " Crore "
        num = num Mod 10000000
    End If

    If num >= 100000 Then
        result = result & NumToWords(num \ 100000) & " Lakh "
        num = num Mod 100000
    End If

    If num >= 1000 Then
        result = result & NumToWords(num \ 1000) & " Thousand "
        num = num Mod 1000
    End If

    If num >= 100 Then
        result = result & ones(num \ 100) & " Hundred "
        num = num Mod 100
    End If

    If num >= 20 Then
        result = result & tens(num \ 10) & " "
        num = num Mod 10
    End If

    If num > 0 Then
        result = result & ones(num) & " "
    End If

    NumToWords = Trim(result)
End Function

Private Function ConvertToWords_Helper(ByVal num As Long) As String
    Dim ones() As String
    Dim tens() As String
    Dim result As String
    
    ones = Array("", "One", "Two", "Three", "Four", "Five", "Six", "Seven", _
                 "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", _
                 "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen")
    tens = Array("", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety")
    
    result = ""
    
    If num >= 10000000 Then ' Crore
        result = result & ConvertToWords_Helper(num \ 10000000) & " Crore "
        num = num Mod 10000000
    End If
    
    If num >= 100000 Then ' Lakh
        result = result & ConvertToWords_Helper(num \ 100000) & " Lakh "
        num = num Mod 100000
    End If
    
    If num >= 1000 Then ' Thousand
        result = result & ConvertToWords_Helper(num \ 1000) & " Thousand "
        num = num Mod 1000
    End If
    
    If num >= 100 Then ' Hundred
        result = result & ones(num \ 100) & " Hundred "
        num = num Mod 100
    End If
    
    If num >= 20 Then
        result = result & tens(num \ 10) & " "
        num = num Mod 10
    End If
    
    If num > 0 Then
        result = result & ones(num) & " "
    End If
    
    ConvertToWords_Helper = Trim(result)
End Function

'------------------------------------------------------------------------------
' 9. PASSWORD GENERATOR
' Usage: =XL_GeneratePassword(8)  <- length parameter
' Example: 12 -> "aB3#xK9@mL2!"
' Note: Volatile function - recalculates on every sheet change
'------------------------------------------------------------------------------
Function XL_GeneratePassword(length As Integer) As String
    Dim chars As String
    Dim result As String
    Dim i As Integer
    Dim charLen As Integer
    
    ' Character pool: uppercase, lowercase, numbers, symbols
    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*"
    charLen = Len(chars)
    
    ' Validate length
    If length < 1 Then length = 8
    If length > 128 Then length = 128
    
    result = ""
    Randomize
    
    ' Ensure at least one of each type for passwords >= 4 chars
    If length >= 4 Then
        result = result & Mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ", Int(Rnd() * 26) + 1, 1) ' uppercase
        result = result & Mid("abcdefghijklmnopqrstuvwxyz", Int(Rnd() * 26) + 1, 1) ' lowercase
        result = result & Mid("0123456789", Int(Rnd() * 10) + 1, 1)                  ' digit
        result = result & Mid("!@#$%^&*", Int(Rnd() * 8) + 1, 1)                    ' symbol
    End If
    
    ' Fill remaining length
    For i = Len(result) + 1 To length
        result = result & Mid(chars, Int(Rnd() * charLen) + 1, 1)
    Next i
    
    ' Shuffle the result (Fisher-Yates)
    Dim arr() As String
    ReDim arr(1 To Len(result))
    For i = 1 To Len(result)
        arr(i) = Mid(result, i, 1)
    Next i
    
    Dim j As Integer
    Dim temp As String
    For i = Len(result) To 2 Step -1
        j = Int(Rnd() * i) + 1
        temp = arr(i)
        arr(i) = arr(j)
        arr(j) = temp
    Next i
    
    result = ""
    For i = 1 To UBound(arr)
        result = result & arr(i)
    Next i
    
    XL_GeneratePassword = result
End Function

'------------------------------------------------------------------------------
' 10. ADD SPACE (Fix CamelCase / PascalCase)
' Usage: =XL_AddSpace(A1)
' Example: "PowerPoint" -> "Power Point"
' Example: "MicrosoftExcel" -> "Microsoft Excel"
'------------------------------------------------------------------------------
Function XL_AddSpace(cell As Range) As String
    Dim s As String
    Dim result As String
    Dim i As Integer
    Dim c As String
    Dim prevC As String
    
    s = CStr(cell.Value)
    
    If Len(s) = 0 Then
        XL_AddSpace = ""
        Exit Function
    End If
    
    result = left(s, 1)
    
    For i = 2 To Len(s)
        c = Mid(s, i, 1)
        prevC = Mid(s, i - 1, 1)
        
        ' Insert space before uppercase letter if previous was lowercase
        If c >= "A" And c <= "Z" Then
            If prevC >= "a" And prevC <= "z" Then
                result = result & " " & c
            ElseIf i < Len(s) Then
                ' Handle consecutive caps like "XMLParser" -> "XML Parser"
                Dim nextC As String
                nextC = Mid(s, i + 1, 1)
                If (nextC >= "a" And nextC <= "z") And (prevC >= "A" And prevC <= "Z") Then
                    result = result & " " & c
                Else
                    result = result & c
                End If
            Else
                result = result & c
            End If
        Else
            result = result & c
        End If
    Next i
    
    ' Clean multiple spaces
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    XL_AddSpace = Trim(result)
End Function

'------------------------------------------------------------------------------
' 11. REVERSE NAME (Last First -> First Last)
' Usage: =XL_ReverseName(A1)
' Example: "Sharma Rahul" -> "Rahul Sharma"
' Example: "Man Iron" -> "Iron Man"
'------------------------------------------------------------------------------
Function XL_ReverseName(cell As Range) As String
    Dim s As String
    Dim parts() As String
    Dim i As Integer
    Dim result As String
    
    s = Trim(CStr(cell.Value))
    
    If Len(s) = 0 Then
        XL_ReverseName = ""
        Exit Function
    End If
    
    ' Remove extra spaces
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    
    parts = Split(s, " ")
    
    If UBound(parts) = 0 Then
        ' Only one word, return as-is
        XL_ReverseName = s
        Exit Function
    End If
    
    ' Reverse the parts array
    Dim left As Integer
    Dim right As Integer
    Dim temp As String
    left = 0
    right = UBound(parts)
    
    Do While left < right
        temp = parts(left)
        parts(left) = parts(right)
        parts(right) = temp
        left = left + 1
        right = right - 1
    Loop
    
    result = ""
    For i = 0 To UBound(parts)
        If result = "" Then
            result = parts(i)
        Else
            result = result & " " & parts(i)
        End If
    Next i
    
    XL_ReverseName = result
End Function

'==============================================================================
' RIBBON BUTTON HANDLERS
' These Sub routines are called from the ribbon buttons.
' They apply the formula to the currently selected cell(s).
'==============================================================================

Sub ExtractText(control As IRibbonControl)
    Call ApplyFormulaToSelection("XL_ExtractText")
End Sub

Sub ExtractNumbers(control As IRibbonControl)
    Call ApplyFormulaToSelection("XL_ExtractNumbers")
End Sub

Sub ExtractEmails(control As IRibbonControl)
    Call ApplyFormulaToSelection("XL_ExtractEmails")
End Sub

Sub MobileNumberOnly(control As IRibbonControl)
    Call ApplyFormulaToSelection("XL_MobileNumberOnly")
End Sub

Sub RemoveSpecialChars(control As IRibbonControl)
    Call ApplyFormulaToSelection("XL_RemoveSpecialChars")
End Sub

Sub ExtractURL(control As IRibbonControl)
    Call ApplyFormulaToSelection("XL_ExtractURL")
End Sub

Sub AgeFromDOB(control As IRibbonControl)
    Call ApplyFormulaToSelection("XL_AgeFromDOB")
End Sub

Sub ConvertToWords(control As IRibbonControl)
    Call ApplyFormulaToSelection("XL_ConvertToWords")
End Sub

Sub GeneratePassword(control As IRibbonControl)
    ' Password needs a length input
    Dim length As Integer
    Dim userInput As String
    
    userInput = InputBox("Enter password length (6-64):", "Generate Password", "12")
    
    If userInput = "" Then Exit Sub
    If Not IsNumeric(userInput) Then
        MsgBox "Please enter a valid number.", vbExclamation
        Exit Sub
    End If
    
    length = CInt(userInput)
    If length < 6 Then length = 6
    If length > 64 Then length = 64
    
    If Selection Is Nothing Then Exit Sub
    
    Dim cell As Range
    For Each cell In Selection
        If Not cell.HasFormula Then
            cell.Value = XL_GeneratePassword(length)
        End If
    Next cell
End Sub

Sub AddSpaceFormula(control As IRibbonControl)
    Call ApplyFormulaToSelection("XL_AddSpace")
End Sub

Sub ReverseName(control As IRibbonControl)
    Call ApplyFormulaToSelection("XL_ReverseName")
End Sub

'------------------------------------------------------------------------------
' HELPER: Apply a formula to selected cells
' Puts formula in the next column if current cell has data,
' OR puts formula directly if cell is empty (yellow practice cells)
'------------------------------------------------------------------------------
Private Sub ApplyFormulaToSelection(formulaName As String)
    If Selection Is Nothing Then
        MsgBox "Please select a cell first.", vbInformation
        Exit Sub
    End If
    
    Dim cell As Range
    Dim targetCell As Range
    
    For Each cell In Selection
        ' If selected cell is empty, check if it's a formula target cell
        ' Put formula referencing the cell to the left
        If cell.Value = "" Or cell.Value = 0 Then
            ' Target cell is empty - put formula referencing left neighbor
            If cell.Column > 1 Then
                Set targetCell = cell
                targetCell.Formula = "=" & formulaName & "(" & cell.Offset(0, -1).Address(False, False) & ")"
            End If
        Else
            ' Selected cell has data - put formula in next column
            Set targetCell = cell.Offset(0, 1)
            targetCell.Formula = "=" & formulaName & "(" & cell.Address(False, False) & ")"
        End If
    Next cell
End Sub

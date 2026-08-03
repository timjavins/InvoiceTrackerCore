' Normalizes values coming off a vendor worksheet so a column holds one data type.
'
' Vendors send the same column typed inconsistently: most rows arrive as floating point, then
' a few come through as text like "$1,234.56" or with stray spaces. Excel preserves whatever
' each cell was, so the column ends up mixed -- and text that looks like a number sorts wrong,
' fails SUM, and breaks XLOOKUP against a numeric key.
'
' The intended types are: dollar amounts as numbers rounded to two decimals, dates as real
' dates, and everything else -- store numbers, invoice numbers, requisition numbers, PO
' numbers -- as text. Identifiers stay text on purpose: they can carry leading zeros or
' letters, and a number would silently drop a leading zero.

' Parses a dollar amount into a Double rounded to two decimals.
'
' Handles the shapes vendors actually send: "$1,234.56", " 1234.56 ", "(45.00)" for a credit,
' "1.234,56" in a European layout, and a value that is already numeric. Returns 0 for
' anything unparseable, and sets ok to False so a caller can tell 0 from "not a number".
Public Function CoerceMoney(ByVal value As Variant, Optional ByRef ok As Boolean) As Double
    ok = False

    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function

    ' Already a number: just round it.
    If Not VarType(value) = vbString Then
        If VBA.IsNumeric(value) Then
            ok = True
            CoerceMoney = Round2(CDbl(value))
        End If
        Exit Function
    End If

    Dim text As String
    text = Trim$(CStr(value))
    If Len(text) = 0 Then Exit Function

    ' Parentheses or a trailing minus mean a credit.
    Dim isNegative As Boolean
    If InStr(text, "(") > 0 And InStr(text, ")") > 0 Then isNegative = True
    If Right$(text, 1) = "-" Then
        isNegative = True
        text = Left$(text, Len(text) - 1)
    End If

    ' Strip currency symbols, spaces (including the non-breaking space Excel sometimes
    ' inherits from a web export), parentheses, and thousands separators.
    text = Replace(text, "$", "")
    text = Replace(text, Chr$(160), "")
    text = Replace(text, " ", "")
    text = Replace(text, "(", "")
    text = Replace(text, ")", "")
    text = Replace(text, Chr$(9), "")

    If Len(text) = 0 Then Exit Function

    ' Decide which separator is the decimal point. "1.234,56" is European; "1,234.56" is not.
    Dim lastComma As Long, lastDot As Long
    lastComma = InStrRev(text, ",")
    lastDot = InStrRev(text, ".")

    If lastComma > 0 And lastDot > 0 Then
        If lastComma > lastDot Then
            ' European: dots group, comma is the decimal.
            text = Replace(text, ".", "")
            text = Replace(text, ",", ".")
        Else
            text = Replace(text, ",", "")
        End If
    ElseIf lastComma > 0 Then
        ' A single comma with exactly two digits after it is a decimal, not a group separator.
        If Len(text) - lastComma = 2 Then
            text = Replace(text, ",", ".")
        Else
            text = Replace(text, ",", "")
        End If
    End If

    ' Leading + is harmless but not numeric to VBA.
    If Left$(text, 1) = "+" Then text = Mid$(text, 2)

    If Not VBA.IsNumeric(text) Then Exit Function

    Dim result As Double
    result = CDbl(text)
    If isNegative Then result = -result

    ok = True
    CoerceMoney = Round2(result)
End Function

' Rounds to two decimals away from zero, which is what money expects. VBA's Round uses
' banker's rounding, so 2.345 would go to 2.34 rather than 2.35.
Public Function Round2(ByVal value As Double) As Double
    Dim scaled As Double
    scaled = value * 100

    If scaled >= 0 Then
        Round2 = Int(scaled + 0.5) / 100
    Else
        Round2 = -Int(-scaled + 0.5) / 100
    End If
End Function

' Parses a date, whether it arrives as a real date, an Excel serial number, or text.
' Returns 0 and sets ok to False when the value is not a date.
Public Function CoerceDate(ByVal value As Variant, Optional ByRef ok As Boolean) As Date
    ok = False

    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function

    If VarType(value) = vbDate Then
        ok = True
        CoerceDate = CDate(value)
        Exit Function
    End If

    ' An Excel serial. Excel's epoch starts at 1; anything below that is not a real date, and
    ' a bare small number is far more likely to be something else entirely.
    If VBA.IsNumeric(value) Then
        Dim serial As Double
        serial = CDbl(value)
        If serial >= 1 And serial < 2958466 Then   ' 2958465 = 9999-12-31
            ok = True
            CoerceDate = CDate(serial)
        End If
        Exit Function
    End If

    Dim text As String
    text = Trim$(CStr(value))
    If Len(text) = 0 Then Exit Function

    On Error Resume Next
    Dim parsed As Date
    parsed = CDate(text)
    If Err.Number = 0 Then
        ok = True
        CoerceDate = parsed
    End If
    On Error GoTo 0
End Function

' Normalizes an identifier -- invoice #, requisition #, PO #, store # -- to trimmed text.
'
' Identifiers are text on purpose. A value like "0148" or "00123456" must keep its leading
' zeros, and Excel would drop them if the cell were numeric. Numeric input that Excel has
' already turned into a float is rendered without the spurious ".0" or scientific notation
' that CStr would produce on a large number.
Public Function CoerceIdentifier(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function

    If VarType(value) = vbString Then
        CoerceIdentifier = Trim$(CStr(value))
        Exit Function
    End If

    If VBA.IsNumeric(value) Then
        Dim d As Double
        d = CDbl(value)

        ' Whole numbers render as integers, so 123456789 does not come back as 1.23457E+08.
        If d = Int(d) And Abs(d) < 1E+15 Then
            CoerceIdentifier = Format$(d, "0")
        Else
            CoerceIdentifier = Trim$(CStr(value))
        End If
        Exit Function
    End If

    CoerceIdentifier = Trim$(CStr(value))
End Function

' Writes a money value to a cell as a number with two-decimal currency formatting, so the
' column is uniformly numeric regardless of how the source cell was typed.
Public Sub WriteMoney(ByVal target As Range, ByVal value As Variant)
    Dim ok As Boolean
    Dim amount As Double
    amount = CoerceMoney(value, ok)

    target.NumberFormat = "#,##0.00"

    If ok Then
        target.value = amount
    Else
        ' Not a number: leave the cell empty rather than write text into a numeric column,
        ' which is what made the column mixed in the first place.
        target.ClearContents
        Debug.Print "WriteMoney: could not read '" & CStr(value) & "' as an amount at " & _
                    target.Address(False, False) & " -- left blank."
    End If
End Sub

' Writes a date value to a cell as a real date.
Public Sub WriteDate(ByVal target As Range, ByVal value As Variant)
    Dim ok As Boolean
    Dim d As Date
    d = CoerceDate(value, ok)

    target.NumberFormat = "mm/dd/yyyy"

    If ok Then
        target.value = d
    Else
        target.ClearContents
        Debug.Print "WriteDate: could not read '" & CStr(value) & "' as a date at " & _
                    target.Address(False, False) & " -- left blank."
    End If
End Sub

' Writes an identifier to a cell as text, preserving leading zeros.
Public Sub WriteIdentifier(ByVal target As Range, ByVal value As Variant)
    target.NumberFormat = "@"
    target.value = CoerceIdentifier(value)
End Sub

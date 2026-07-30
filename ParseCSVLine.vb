' Splits one CSV line into fields, honouring quoted fields (RFC 4180).
'
' Both variants got this wrong in different ways, and both corrupted real Coupa exports:
'
'   Securitas set inQuotes = True on a quote but never back to False, so everything after
'   the first quoted field collapsed into a single column:
'       "Smith, John",100,ok   ->   ["Smith, John,100,ok"]
'
'   JCI split on every comma, so a quoted field containing a comma broke apart:
'       "Smith, John",100,ok   ->   ['"Smith', ' John"', '100', 'ok']
'
' This version toggles quote state, treats a doubled quote inside a quoted field as one
' literal quote, and returns a 0-based array (callers index with headerIndex - 1).
Public Function ParseCSVLine(ByVal lineData As String) As String()
    Dim values() As String
    Dim valueCount As Long
    Dim currentValue As String
    Dim inQuotes As Boolean
    Dim ch As String
    Dim i As Long

    ReDim values(0 To 0)

    i = 1
    Do While i <= Len(lineData)
        ch = Mid$(lineData, i, 1)

        If ch = """" Then
            If inQuotes And i < Len(lineData) Then
                If Mid$(lineData, i + 1, 1) = """" Then
                    ' Escaped quote inside a quoted field.
                    currentValue = currentValue & """"
                    i = i + 2
                    GoTo ContinueLoop
                End If
            End If
            inQuotes = Not inQuotes

        ElseIf ch = "," And Not inQuotes Then
            values(valueCount) = currentValue
            valueCount = valueCount + 1
            ReDim Preserve values(0 To valueCount)
            currentValue = vbNullString

        Else
            currentValue = currentValue & ch
        End If

        i = i + 1
ContinueLoop:
    Loop

    values(valueCount) = currentValue
    ParseCSVLine = values
End Function

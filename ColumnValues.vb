' Safe reads of worksheet data into VBA values.
'
' Both helpers exist because Excel hands back two shapes that break naive code, and every
' array-processing routine in this project had to guard against them separately -- or forgot to.

' A worksheet column as an array that is always two-dimensional.
'
' Range.Value returns a bare scalar when the range is a single cell, so UBound(v, 1) raises a type
' mismatch on a sheet with exactly one data row. ConvertStoreNumbersOn guarded this inline;
' UpdateSearchValues and GenerateAccrualsReport did not, and both crashed on a one-row tracker.
'
' Returns Empty when the range would be invalid, so callers should check lastRow themselves before
' relying on a result.
Public Function ColumnValues(ByVal ws As Worksheet, _
                             ByVal colLetter As String, _
                             ByVal firstRow As Long, _
                             ByVal lastRow As Long) As Variant
    If ws Is Nothing Then Exit Function
    If lastRow < firstRow Then Exit Function

    Dim raw As Variant
    raw = ws.Range(colLetter & firstRow & ":" & colLetter & lastRow).Value

    If IsArray(raw) Then
        ColumnValues = raw
        Exit Function
    End If

    Dim single_ As Variant
    ReDim single_(1 To 1, 1 To 1)
    single_(1, 1) = raw
    ColumnValues = single_
End Function

' A cell value as trimmed text, rendering Excel error values as "".
'
' A Variant holding an error subtype raises a type mismatch on almost any use, including a plain
' comparison to a number: `If colB(i, 1) = 200` blows up when column B holds #N/A. Formula columns
' produce #N/A routinely -- any store number missing from 'BU List' puts one in column B -- so
' every comparison against a formula column has to come through here.
Public Function SafeText(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    If IsEmpty(value) Then Exit Function
    If IsNull(value) Then Exit Function
    SafeText = VBA.Trim$(CStr(value))
End Function

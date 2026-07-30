' Resolves a set of column headers to their column indexes on a worksheet, so callers can
' work by header name instead of hardcoding letters.
'
' Matching is case-insensitive and ignores surrounding whitespace. Headers that are not
' found are simply absent from the result, so callers should check .Exists before reading.
'
' Note: JCI shipped this same function in a file named GetColumnHeaderIndexes.vb (word
' order reversed). Core standardises on the filename matching the function name; the
' variant file was removed.
Public Function GetHeaderColumnIndexes(ByVal ws As Worksheet, _
                                       ByVal headerRow As Long, _
                                       ByVal requiredCols As Variant) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    Set GetHeaderColumnIndexes = result

    If ws Is Nothing Then Exit Function
    If headerRow < 1 Then Exit Function

    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    ' Index the sheet's headers once, then look up each requested name.
    Dim present As Object
    Set present = CreateObject("Scripting.Dictionary")

    Dim c As Long
    Dim key As String
    For c = 1 To lastCol
        key = VBA.UCase$(VBA.Trim$(CStr(ws.Cells(headerRow, c).Value)))
        If Len(key) > 0 Then
            If Not present.Exists(key) Then present.Add key, c
        End If
    Next c

    Dim i As Long
    Dim wanted As String
    For i = LBound(requiredCols) To UBound(requiredCols)
        wanted = VBA.UCase$(VBA.Trim$(CStr(requiredCols(i))))
        If present.Exists(wanted) Then
            If Not result.Exists(CStr(requiredCols(i))) Then
                result.Add CStr(requiredCols(i)), present(wanted)
            End If
        End If
    Next i
End Function

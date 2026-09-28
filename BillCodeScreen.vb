' Decides, before anything is written, which incoming bill rows to skip.
'
' Two different tests, because Securitas monitoring bills put one bill code on several store
' rows and repair/installation bills never do (see the batch-tracking spec, section 2):
'   already-processed  the code is already in the tracker (the rows BEFORE this import)
'   repeat-in-file     the code appears on more than one incoming row, and the caller said
'                      repeats are not allowed for this bill type
' The old check ran after the rows were written and counted the whole column, new rows
' included, so every monitoring row after the first looked like a duplicate of its siblings.
'
' Knows nothing about bill types: the caller decides allowRepeatsInFile. Pure over arrays,
' so the core harness can test it without a tracker.
'
' Every row of a repeated code is flagged, not just the later ones: the tool cannot tell
' which copy is right, so a person decides.
Public Function ClassifyBillCodes(ByVal existingCodes As Variant, ByVal incomingCodes As Variant, _
                                  ByVal allowRepeatsInFile As Boolean) As Variant
    Dim existing As Object
    Dim counts As Object
    Dim result() As Variant
    Dim i As Long
    Dim k As String

    If Not IsArray(incomingCodes) Then ClassifyBillCodes = Array(): Exit Function
    If UBound(incomingCodes) < LBound(incomingCodes) Then ClassifyBillCodes = Array(): Exit Function

    Set existing = CreateObject("Scripting.Dictionary")
    existing.CompareMode = vbTextCompare
    If IsArray(existingCodes) Then
        For i = LBound(existingCodes) To UBound(existingCodes)
            k = BcsKey(existingCodes(i))
            If Len(k) > 0 Then existing(k) = True
        Next i
    End If

    Set counts = CreateObject("Scripting.Dictionary")
    counts.CompareMode = vbTextCompare
    For i = LBound(incomingCodes) To UBound(incomingCodes)
        k = BcsKey(incomingCodes(i))
        If Len(k) > 0 Then counts(k) = counts(k) + 1
    Next i

    ReDim result(LBound(incomingCodes) To UBound(incomingCodes))
    For i = LBound(incomingCodes) To UBound(incomingCodes)
        k = BcsKey(incomingCodes(i))
        If Len(k) = 0 Then
            result(i) = ""
        ElseIf existing.Exists(k) Then
            result(i) = "already-processed"
        ElseIf Not allowRepeatsInFile And counts(k) > 1 Then
            result(i) = "repeat-in-file"
        Else
            result(i) = ""
        End If
    Next i

    ClassifyBillCodes = result
End Function

' One column of a sheet as a 0-based 1-D array. Range.Value returns a bare value, not an array,
' for a single cell, which is the case this exists to absorb.
Public Function BillCodeColumnValues(ByVal ws As Worksheet, ByVal col As Long, _
                                     ByVal firstRow As Long, ByVal lastRow As Long) As Variant
    Dim raw As Variant
    Dim result() As Variant
    Dim i As Long

    If lastRow < firstRow Then BillCodeColumnValues = Array(): Exit Function

    raw = ws.Range(ws.Cells(firstRow, col), ws.Cells(lastRow, col)).Value
    ReDim result(0 To lastRow - firstRow)
    If IsArray(raw) Then
        For i = 0 To lastRow - firstRow
            result(i) = raw(i + 1, 1)
        Next i
    Else
        result(0) = raw
    End If
    BillCodeColumnValues = result
End Function

Private Function BcsKey(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    BcsKey = Trim$(CStr(v))
End Function

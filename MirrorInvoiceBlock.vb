' Replaces dst's owned block with a values-only copy of src, and returns the rows copied, or
' -1 if it wrote nothing.
'
' Built for the All-Years archive (batch-tracking spec, section 3): src is the working book's
' Invoices sheet, dst the archive's, which has the same columns plus a trailing "Source File"
' tag. Rows whose tag is in ownedTags belong to the mirror and are replaced wholesale on every
' run; every other row is history and is never touched.
'
' Whole-block rather than row matching, because no row key is unique: a monitoring bill puts
' one BILL CODE on several store rows in both workbooks. Replacing the block needs no key,
' and re-running it is always safe.
'
' Owned rows are deleted run by run and the new block appended, rather than reading the whole
' sheet and writing it back, because writing a legacy text value like "0175" back into a
' General cell would silently turn it into the number 175.
Public Function MirrorInvoiceBlock(ByVal src As Worksheet, ByVal dst As Worksheet, _
                                   ByVal writeTag As String, ByVal ownedTags As Variant) As Long
    Dim srcCols As Long, tagCol As Long, keyCol As Long, c As Long
    Dim srcLast As Long, dstLast As Long, n As Long, r As Long, runEnd As Long, newFirst As Long
    Dim tags As Variant, data As Variant, one() As Variant

    MirrorInvoiceBlock = -1
    On Error GoTo Failed

    ' 1. Headers must line up exactly.
    srcCols = src.Cells(1, src.Columns.Count).End(xlToLeft).Column
    tagCol = srcCols + 1
    For c = 1 To srcCols
        If MibHeader(src.Cells(1, c).Value) <> MibHeader(dst.Cells(1, c).Value) Then Exit Function
        If MibHeader(src.Cells(1, c).Value) = "BILL CODE" Then keyCol = c
    Next c
    If keyCol = 0 Then Exit Function
    If MibHeader(dst.Cells(1, tagCol).Value) <> "SOURCE FILE" Then Exit Function
    If dst.Cells(1, dst.Columns.Count).End(xlToLeft).Column <> tagCol Then Exit Function

    ' 2. Real last rows, never UsedRange (the archive's is stale from old deletions).
    srcLast = MibLastRow(src, 1, keyCol)
    If srcLast < 2 Then Exit Function          ' an empty source must not wipe the owned block
    n = srcLast - 1
    dstLast = MibLastRow(dst, 1, tagCol)

    ' 3. Delete owned rows, bottom-up, one delete per contiguous run. Tags are read once up
    '    front; deleting below r never shifts rows at or above r, so the indexes stay valid.
    If dstLast >= 2 Then
        tags = dst.Range(dst.Cells(2, tagCol), dst.Cells(dstLast, tagCol)).Value
        r = dstLast
        Do While r >= 2
            If MibOwned(MibAt(tags, r - 1), ownedTags) Then
                runEnd = r
                Do While r >= 2
                    If Not MibOwned(MibAt(tags, r - 1), ownedTags) Then Exit Do
                    r = r - 1
                Loop
                dst.Rows((r + 1) & ":" & runEnd).Delete
            Else
                r = r - 1
            End If
        Loop
    End If

    ' 4. Append src as values. Formats first, so text stays text when the values land.
    newFirst = MibLastRow(dst, 1, tagCol) + 1
    If newFirst < 2 Then newFirst = 2
    data = src.Range(src.Cells(2, 1), src.Cells(srcLast, srcCols)).Value
    ' A 1x1 range returns a bare value, not an array.
    If Not IsArray(data) Then ReDim one(1 To 1, 1 To 1): one(1, 1) = data: data = one
    For c = 1 To srcCols
        dst.Range(dst.Cells(newFirst, c), dst.Cells(newFirst + n - 1, c)).NumberFormat = _
            MibFormatFor(data, c, src.Cells(2, c).NumberFormat)
    Next c
    dst.Range(dst.Cells(newFirst, 1), dst.Cells(newFirst + n - 1, srcCols)).Value = data
    dst.Range(dst.Cells(newFirst, tagCol), dst.Cells(newFirst + n - 1, tagCol)).Value = writeTag

    MirrorInvoiceBlock = n
    Exit Function

Failed:
    MirrorInvoiceBlock = -1
End Function

Private Function MibHeader(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    MibHeader = UCase$(Trim$(CStr(v)))
End Function

Private Function MibLastRow(ByVal ws As Worksheet, ByVal colA As Long, ByVal colB As Long) As Long
    Dim a As Long, b As Long
    a = ws.Cells(ws.Rows.Count, colA).End(xlUp).Row
    b = ws.Cells(ws.Rows.Count, colB).End(xlUp).Row
    MibLastRow = IIf(a > b, a, b)
End Function

Private Function MibAt(ByVal tags As Variant, ByVal i As Long) As Variant
    If IsArray(tags) Then MibAt = tags(i, 1) Else MibAt = tags
End Function

Private Function MibOwned(ByVal tag As Variant, ByVal ownedTags As Variant) As Boolean
    Dim i As Long
    If IsError(tag) Or IsEmpty(tag) Then Exit Function
    For i = LBound(ownedTags) To UBound(ownedTags)
        If StrComp(Trim$(CStr(tag)), Trim$(CStr(ownedTags(i))), vbTextCompare) = 0 Then
            MibOwned = True
            Exit Function
        End If
    Next i
End Function

' A column holding any numeric-looking text ("0003", "6005463095") is forced to Text, or the
' value write converts it to a number. Otherwise the source's own format carries over.
Private Function MibFormatFor(ByVal data As Variant, ByVal c As Long, ByVal fallback As String) As String
    Dim i As Long
    For i = LBound(data, 1) To UBound(data, 1)
        If VarType(data(i, c)) = vbString Then
            If Len(data(i, c)) > 0 And IsNumeric(data(i, c)) Then
                MibFormatFor = "@"
                Exit Function
            End If
        End If
    Next i
    MibFormatFor = fallback
End Function

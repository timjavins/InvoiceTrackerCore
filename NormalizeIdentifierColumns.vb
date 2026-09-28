' Normalizes the tracker's identifier columns -- REQ #, invoice numbers, payment # -- to text,
' the same way ConvertStoreNumbers.vb normalizes STORE #. See
' notes/identifier-types-investigation.md.
'
' Why this exists: every Coupa-side sheet (Coupa Reqs, Coupa Invs, Coupa POs) stores its key
' column as text -- confirmed on both live workbooks. XLOOKUP/VLOOKUP/Match never coerce a
' numeric key to match a text value or vice versa, so a tracker cell that lands as a Number
' (Excel re-numericizes a numeric-looking string typed, pasted, or written into a General cell)
' silently breaks every formula keyed on it. This sweeps the tracker's own identifier columns
' back to text on every Refresh, rather than relying on a one-time migration to stay clean.

' Pure conversion: an identifier value as trimmed text, preserving digits exactly (no
' scientific notation, no dropped leading character) and leaving non-identifier shapes alone.
'
'   - A whole-number numeric value (Double/Long/Integer/Currency/Decimal, etc.) becomes
'     Format$(v, "0") -- so 8059053607 becomes "8059053607", never "8.05905E+09".
'   - A non-whole-number numeric value becomes CStr(v) -- an identifier is never fractional in
'     practice, but this still returns a legible string rather than guessing at rounding.
'   - A string becomes Trim$ of itself.
'   - Empty, Null, and error values pass through unchanged -- there is nothing to normalize,
'     and guessing a replacement for any of them would destroy information (a genuine blank,
'     a #N/A propagated from elsewhere, or a deliberately unset cell).
Public Function NormalizeIdentifier(ByVal v As Variant) As Variant
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        NormalizeIdentifier = v
        Exit Function
    End If

    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte, vbLongLong
            Dim d As Double
            d = CDbl(v)
            If d = Int(d) Then
                NormalizeIdentifier = Format$(v, "0")
            Else
                NormalizeIdentifier = CStr(v)
            End If

        Case vbString
            NormalizeIdentifier = Trim$(CStr(v))

        Case Else
            ' Boolean, Date, Object, etc. -- not a shape this column should ever hold; handed
            ' back unchanged rather than guessed at.
            NormalizeIdentifier = v
    End Select
End Function

' Sweeps one or more columns on ws, rewriting every non-formula cell that is not already
' normalized text. cols is an array of column letters (already resolved through
' TenantColLetter by the caller below), so this is testable against any sheet/column shape
' without a tenant workbook.
'
' Returns the count of cells actually changed. Never raises and never shows a dialog -- this
' runs from Refresh, where an unhandled error must not abort the rest of the run, and the test
' harness cannot survive a modal dialog. On error, whatever was already normalized stays
' normalized; the count reflects that partial progress rather than being lost.
Public Function NormalizeIdentifierColumnsOn(ByVal ws As Worksheet, ByVal cols As Variant) As Long
    Dim changed As Long
    changed = 0

    On Error GoTo Done

    If ws Is Nothing Then GoTo Done
    If Not IsArray(cols) Then GoTo Done

    ' UsedRange's extent is unaffected by AutoFilter-hidden rows -- only by actual content --
    ' unlike End(xlUp) or Find, which both walk visible cells only and so would silently miss a
    ' filtered-out trailing row. See MirrorInvoiceBlock.vb for the same trap measured directly.
    Dim usedLastRow As Long
    usedLastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

    Dim c As Long
    For c = LBound(cols) To UBound(cols)
        Dim colLetter As String
        colLetter = Trim$(CStr(cols(c)))

        If Len(colLetter) > 0 Then
            Dim colLastRow As Long
            colLastRow = LastNonEmptyRow(ws, colLetter, usedLastRow)

            If colLastRow >= 2 Then
                Dim r As Long
                For r = 2 To colLastRow
                    Dim cell As Range
                    Set cell = ws.Cells(r, colLetter)

                    ' Never write to a formula cell -- rewriting its value replaces the
                    ' formula with a literal.
                    If Not cell.HasFormula Then
                        Dim raw As Variant
                        raw = cell.Value

                        If IsError(raw) Or IsNull(raw) Or IsEmpty(raw) Then
                            ' Nothing to normalize.

                        ElseIf VarType(raw) = vbString Then
                            Dim trimmed As String
                            trimmed = CStr(NormalizeIdentifier(raw))
                            If trimmed <> CStr(raw) Then
                                cell.NumberFormat = "@"
                                cell.Value = trimmed
                                changed = changed + 1
                            End If
                            ' Else: already normalized text -- left untouched on purpose, to
                            ' keep co-author writes minimal.

                        Else
                            ' Numeric (or any other non-blank, non-error, non-string type):
                            ' always needs to become text, regardless of whether the digits
                            ' themselves change.
                            cell.NumberFormat = "@"
                            cell.Value = CStr(NormalizeIdentifier(raw))
                            changed = changed + 1
                        End If
                    End If
                Next r
            End If
        End If
    Next c

Done:
    NormalizeIdentifierColumnsOn = changed
End Function

' Last non-blank row in one column, immune to filter-hidden rows -- see the comment above on
' why UsedRange is used instead of End(xlUp)/Find. Reads the column once into memory and scans
' it in VBA rather than touching the sheet cell by cell.
Private Function LastNonEmptyRow(ByVal ws As Worksheet, ByVal colLetter As String, _
                                 ByVal usedLastRow As Long) As Long
    If usedLastRow < 2 Then Exit Function

    Dim values As Variant
    values = ws.Range(colLetter & "2:" & colLetter & usedLastRow).Value

    ' A one-row range comes back as a bare value rather than a 2D array.
    If Not IsArray(values) Then
        If Not (IsEmpty(values) Or IsNull(values) Or IsError(values)) Then
            If Len(Trim$(CStr(values))) > 0 Then LastNonEmptyRow = 2
        End If
        Exit Function
    End If

    Dim i As Long
    Dim v As Variant
    For i = UBound(values, 1) To LBound(values, 1) Step -1
        v = values(i, 1)
        If Not (IsEmpty(v) Or IsNull(v) Or IsError(v)) Then
            If Len(Trim$(CStr(v))) > 0 Then
                LastNonEmptyRow = i + 1   ' array row 1 == sheet row 2
                Exit Function
            End If
        End If
    Next i
End Function

' Normalizes the tenant's declared identifier columns on its tracker sheet.
'
' Silent by default -- called from Refresh, where a modal dialog per step is noise. Pass
' announce:=True for interactive use.
'
' manageProtection follows ConvertStoreNumbers's convention exactly: default True because some
' callers protect the sheet before calling and rely on each step unprotecting itself; callers
' that already unprotect around a whole sequence pass False to avoid redundant toggling.
'
' TenantIdentifierColumns() returning an empty array is a no-op -- a tenant that has not (yet)
' declared any identifier columns simply gets nothing done here.
Public Sub NormalizeIdentifierColumns(Optional ByVal announce As Boolean = False, _
                                      Optional ByVal manageProtection As Boolean = True)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TenantSheetName("tracker"))

    Dim concepts As Variant
    concepts = TenantIdentifierColumns()

    If Not IsArray(concepts) Then Exit Sub
    If UBound(concepts) < LBound(concepts) Then Exit Sub

    Dim cols() As String
    ReDim cols(LBound(concepts) To UBound(concepts))

    Dim i As Long
    For i = LBound(concepts) To UBound(concepts)
        cols(i) = TenantColLetter(CStr(concepts(i)))
    Next i

    If manageProtection Then
        ws.Activate
        UnprotectSheet
    End If

    Dim changed As Long
    changed = NormalizeIdentifierColumnsOn(ws, cols)

    If manageProtection Then ProtectSheet

    If announce Then
        MsgBox changed & " identifier value(s) normalized to text.", vbInformation
    End If
End Sub

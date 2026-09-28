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
' Each column is read in bulk -- one .Value read and one .Formula read over rows 2..lastRow --
' and computed entirely in memory before anything is written back, the same read/compute/write
' shape ConvertStoreNumbers.vb already uses. That turns a ~3500-row column from ~7000 COM round
' trips (HasFormula + Value per cell) into two reads and, in the common case where the column
' has no formula cells, exactly one bulk write. See NicSweepColumn.
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

    ' Reads are filter-safe (LastNonEmptyRow above, and the bulk .Value/.Formula reads in
    ' NicSweepColumn, both read every physical row regardless of visibility). Writes are not:
    ' assigning a multi-row array to a Range while FilterMode is True does not reliably map
    ' array rows to sheet rows -- confirmed directly (see notes/normalize-bulk-report.md) --
    ' hidden rows silently keep their old value while visible rows can receive the WRONG
    ' element instead of their own. So whenever the sheet has an active filter, every write
    ' below goes one row at a time instead of in a bulk range, which is exactly what this
    ' module did before this rewrite and so is already proven correct under a filter.
    Dim filterActive As Boolean
    filterActive = ws.FilterMode

    Dim c As Long
    For c = LBound(cols) To UBound(cols)
        Dim colLetter As String
        colLetter = Trim$(CStr(cols(c)))

        If Len(colLetter) > 0 Then
            Dim colLastRow As Long
            colLastRow = LastNonEmptyRow(ws, colLetter, usedLastRow)

            If colLastRow >= 2 Then
                changed = changed + NicSweepColumn(ws, colLetter, colLastRow, filterActive)
            End If
        End If
    Next c

Done:
    NormalizeIdentifierColumnsOn = changed
End Function

' Bulk read/compute/write for one column, rows 2..colLastRow. Reads .Value and .Formula once
' each over the whole range, decides in memory which rows are formula cells (never touched,
' per guarantee #1 -- rewriting a formula cell's value would replace the formula with a
' literal) and which non-formula cells actually need a change, then writes back with as few
' COM calls as possible:
'   - filterActive (the sheet has an active AutoFilter hiding rows right now): a multi-row
'     array write is unsafe in this state -- see NicWritePerCell -- so every changed,
'     non-formula row is written one cell at a time, exactly as this module did before this
'     rewrite. Rare in a scheduled Refresh, but must stay correct if the user is filtering the
'     tracker when it runs.
'   - No formula cells and no active filter (the common, currently-only case for every
'     configured tenant column): one blanket NumberFormat = "@" plus one blanket Value = array
'     over the whole range, exactly ConvertStoreNumbersOn's shape. Cells that don't need a
'     value change keep their own current value in the array, so the bulk write doesn't alter
'     them.
'   - A formula cell mixed into the column, filter not active (rare, never configured today,
'     but must stay correct if it ever happens): the whole-range write is unsafe, since
'     writing .Value to a range containing a formula cell would wipe that formula. Instead,
'     NicWriteRuns writes one bulk NumberFormat/Value per maximal contiguous run of changed,
'     non-formula rows, so a long clean stretch still costs one call and a formula cell is
'     never in any written range.
' Returns the count of cells actually changed (numeric conversions plus string cells whose
' normalized form differs from the raw text).
Private Function NicSweepColumn(ByVal ws As Worksheet, ByVal colLetter As String, _
                                ByVal lastRow As Long, ByVal filterActive As Boolean) As Long
    Dim rowCount As Long
    rowCount = lastRow - 1

    Dim targetRange As Range
    Set targetRange = ws.Range(colLetter & "2:" & colLetter & lastRow)

    Dim rawValues As Variant
    rawValues = targetRange.Value

    Dim rawFormulas As Variant
    rawFormulas = targetRange.Formula

    Dim isFormulaRow() As Boolean
    ReDim isFormulaRow(1 To rowCount)

    Dim needsWrite() As Boolean
    ReDim needsWrite(1 To rowCount)

    Dim targetValues() As Variant
    ReDim targetValues(1 To rowCount, 1 To 1)

    Dim anyFormula As Boolean
    anyFormula = False

    Dim changed As Long
    changed = 0

    Dim r As Long
    For r = 1 To rowCount
        Dim rawValue As Variant
        Dim rawFormula As Variant

        ' A single-row range comes back as a bare scalar rather than a 2D array.
        If rowCount = 1 Then
            rawValue = rawValues
            rawFormula = rawFormulas
        Else
            rawValue = rawValues(r, 1)
            rawFormula = rawFormulas(r, 1)
        End If

        If NicIsFormulaText(rawFormula) Then
            isFormulaRow(r) = True
            anyFormula = True
            targetValues(r, 1) = rawValue

        ElseIf IsError(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then
            ' Nothing to normalize.
            targetValues(r, 1) = rawValue

        ElseIf VarType(rawValue) = vbString Then
            Dim trimmed As String
            trimmed = CStr(NormalizeIdentifier(rawValue))
            If trimmed <> CStr(rawValue) Then
                targetValues(r, 1) = trimmed
                needsWrite(r) = True
                changed = changed + 1
            Else
                ' Already normalized text -- left untouched on purpose, to keep co-author
                ' writes minimal.
                targetValues(r, 1) = rawValue
            End If

        Else
            ' Numeric (or any other non-blank, non-error, non-string type): always needs to
            ' become text, regardless of whether the digits themselves change.
            targetValues(r, 1) = CStr(NormalizeIdentifier(rawValue))
            needsWrite(r) = True
            changed = changed + 1
        End If
    Next r

    If changed = 0 Then
        NicSweepColumn = 0
        Exit Function
    End If

    If filterActive Then
        NicWritePerCell ws, colLetter, isFormulaRow, needsWrite, targetValues, rowCount
    ElseIf Not anyFormula Then
        targetRange.NumberFormat = "@"
        targetRange.Value = targetValues
    Else
        NicWriteRuns ws, colLetter, isFormulaRow, needsWrite, targetValues, rowCount
    End If

    NicSweepColumn = changed
End Function

' True when a bulk-read .Formula value is formula text starting with "=" -- the one way to
' tell a formula cell from a literal without a second trip to the sheet (HasFormula per cell).
' Guards IsError/IsNull first, either of which would make CStr raise.
Private Function NicIsFormulaText(ByVal formulaValue As Variant) As Boolean
    If IsError(formulaValue) Or IsNull(formulaValue) Then Exit Function
    NicIsFormulaText = (Left$(CStr(formulaValue), 1) = "=")
End Function

' Writes each changed, non-formula row with its own single-cell NumberFormat/Value write --
' used only while the sheet has an active AutoFilter (ws.FilterMode). Confirmed directly against
' a live Excel COM host: assigning a multi-row array to a Range that spans an AutoFilter-hidden
' row does not write each row to its own cell -- the hidden row silently keeps its old value
' and every *visible* row in that same write can receive the wrong array element instead of its
' own (observed: every visible row received the array's first element). A single-cell write has
' no such ambiguity, and this is exactly what NormalizeIdentifierColumnsOn did before this
' rewrite, so it is already proven correct under a filter -- see
' notes/normalize-bulk-report.md for the reproduction. rowIsFormula, rowNeedsWrite, and
' targetValues are 1-based over rowCount rows, where array row i is sheet row i + 1.
Private Sub NicWritePerCell(ByVal ws As Worksheet, ByVal colLetter As String, _
                            ByRef rowIsFormula() As Boolean, ByRef rowNeedsWrite() As Boolean, _
                            ByRef targetValues() As Variant, ByVal rowCount As Long)
    Dim r As Long
    For r = 1 To rowCount
        If (Not rowIsFormula(r)) And rowNeedsWrite(r) Then
            Dim cell As Range
            Set cell = ws.Cells(r + 1, colLetter)
            cell.NumberFormat = "@"
            cell.Value = targetValues(r, 1)
        End If
    Next r
End Sub

' Writes only maximal contiguous runs of rows that are both non-formula and actually changed,
' one bulk NumberFormat/Value write per run -- used only when a formula cell is mixed into the
' column, so the whole-range write in NicSweepColumn would otherwise wipe it. rowIsFormula,
' rowNeedsWrite, and targetValues are 1-based over rowCount rows, where array row i is sheet
' row i + 1.
Private Sub NicWriteRuns(ByVal ws As Worksheet, ByVal colLetter As String, _
                         ByRef rowIsFormula() As Boolean, ByRef rowNeedsWrite() As Boolean, _
                         ByRef targetValues() As Variant, ByVal rowCount As Long)
    Dim runStart As Long
    runStart = 0

    Dim r As Long
    For r = 1 To rowCount + 1
        Dim inRun As Boolean
        inRun = False
        If r <= rowCount Then
            inRun = (Not rowIsFormula(r)) And rowNeedsWrite(r)
        End If

        If inRun Then
            If runStart = 0 Then runStart = r
        ElseIf runStart <> 0 Then
            Dim runEnd As Long
            runEnd = r - 1

            Dim runRange As Range
            Set runRange = ws.Range(colLetter & (runStart + 1) & ":" & colLetter & (runEnd + 1))

            Dim runValues() As Variant
            ReDim runValues(1 To runEnd - runStart + 1, 1 To 1)

            Dim j As Long
            For j = runStart To runEnd
                runValues(j - runStart + 1, 1) = targetValues(j, 1)
            Next j

            runRange.NumberFormat = "@"
            runRange.Value = runValues

            runStart = 0
        End If
    Next r
End Sub

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

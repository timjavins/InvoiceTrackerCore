' Finds the real last data row of an imported bill export, excluding any trailing
' summary/total row a vendor's export tool appended (a pivot-style "Grand Total" line, a
' subtotal, etc.) -- WITHOUT assuming that row is blank in any particular column.
'
' Two confirmed shapes already broke a single-column anchor: one Securitas export left
' STORE # blank on its footer and put "Grand Total" in BILL CODE; a later one puts
' "Grand Total" directly in STORE # and leaves BILL CODE (and everything else) blank instead.
' There is no column that is reliably blank on every vendor's summary row, so this judges the
' row's SHAPE instead of trusting any one column's contents. See
' RowValuesLookLikeSummary for the actual signals.

' Whether one already-collected row of values looks like a summary/total row rather than a
' real data row. Three independent, individually weak signals; any two together call it a
' summary row, so a real row is never misclassified just for tripping one of them (being
' sparse, or happening to hold this batch's largest single dollar amount, or -- unlikely, but
' possible in a free-text notes field -- containing the substring "total" somewhere).
'
'   - a cell's whole trimmed text (not a substring match, to avoid a false hit on a
'     TRANSACTION DETAILS note like "replaced totaled panel") is a label like "Total" or
'     "Grand Total"
'   - the row is mostly blank -- a summary row is typically a label and a couple of amount
'     columns, nothing else
'   - the row holds a dollar amount at least as large as the largest amount anywhere else in
'     the scanned block -- true by construction for a genuine grand total (the sum of every
'     other row's amount), and a real single line item essentially never ties or exceeds the
'     sum of everything else
'
' rowValues is 1-based or otherwise contiguous over some LBound..UBound; only the bounds and
' contents matter, not the base.
Public Function RowValuesLookLikeSummary(ByVal rowValues As Variant, _
                                         ByVal maxAmountInBlock As Double) As Boolean
    Dim c As Long
    Dim v As Variant
    Dim text As String
    Dim normalized As String
    Dim blankCount As Long
    Dim totalCols As Long
    Dim hasLabel As Boolean
    Dim hasMoney As Boolean
    Dim hasMaxAmount As Boolean
    Dim amt As Double
    Dim ok As Boolean

    totalCols = UBound(rowValues) - LBound(rowValues) + 1
    If totalCols < 1 Then Exit Function

    For c = LBound(rowValues) To UBound(rowValues)
        v = rowValues(c)

        If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
            blankCount = blankCount + 1
        Else
            text = Trim$(CStr(v))
            If Len(text) = 0 Then
                blankCount = blankCount + 1
            Else
                normalized = UCase$(text)
                Select Case normalized
                    Case "TOTAL", "TOTALS", "GRAND TOTAL", "GRAND TOTALS", _
                         "SUBTOTAL", "SUB TOTAL", "SUBTOTALS"
                        hasLabel = True
                End Select

                amt = CoerceMoney(v, ok)
                If ok And amt > 0 Then
                    hasMoney = True
                    If amt >= maxAmountInBlock Then hasMaxAmount = True
                End If
            End If
        End If
    Next c

    Dim signals As Long
    If hasLabel Then signals = signals + 1
    If blankCount >= (totalCols \ 2) Then signals = signals + 1
    If hasMoney And hasMaxAmount Then signals = signals + 1

    RowValuesLookLikeSummary = (signals >= 2)
End Function

' Worksheet entry point. rawLastRow is a caller's own coarse starting guess (however they
' found it -- End(xlUp) on any column, UsedRange, etc.); this walks upward from there,
' skipping any trailing blank row or one RowValuesLookLikeSummary calls a summary row, and
' returns the first row that is neither. Never returns a row at or above headerRow.
Public Function LastRealDataRow(ByVal ws As Worksheet, ByVal rawLastRow As Long, _
                                ByVal headerRow As Long, ByVal firstCol As Long, _
                                ByVal lastCol As Long) As Long
    If rawLastRow <= headerRow Then
        LastRealDataRow = rawLastRow
        Exit Function
    End If
    If lastCol < firstCol Then
        LastRealDataRow = rawLastRow
        Exit Function
    End If

    Dim maxAmount As Double
    Dim r As Long, c As Long, amt As Double, ok As Boolean
    For r = headerRow + 1 To rawLastRow
        For c = firstCol To lastCol
            amt = CoerceMoney(ws.Cells(r, c).Value, ok)
            If ok And amt > maxAmount Then maxAmount = amt
        Next c
    Next r

    Dim lastRow As Long
    lastRow = rawLastRow
    Do While lastRow > headerRow
        Dim rowValues() As Variant
        ReDim rowValues(firstCol To lastCol)
        Dim rowBlank As Boolean
        rowBlank = True
        For c = firstCol To lastCol
            rowValues(c) = ws.Cells(lastRow, c).Value
            If Not (IsEmpty(rowValues(c)) Or IsError(rowValues(c)) Or _
                    Len(Trim$(CStr(rowValues(c) & ""))) = 0) Then
                rowBlank = False
            End If
        Next c

        If rowBlank Or RowValuesLookLikeSummary(rowValues, maxAmount) Then
            lastRow = lastRow - 1
        Else
            Exit Do
        End If
    Loop

    LastRealDataRow = lastRow
End Function

' Array entry point, for a caller that already loaded its source into memory (e.g. JCI
' AddNewBills via LoadCoupaSourceData) rather than reading a live worksheet. sourceData is a
' 1-based, row-major 2D array. Same contract as LastRealDataRow otherwise.
Public Function LastRealDataRowInArray(ByVal sourceData As Variant, ByVal rawLastRow As Long, _
                                       ByVal headerRow As Long, ByVal firstCol As Long, _
                                       ByVal lastCol As Long) As Long
    If rawLastRow <= headerRow Then
        LastRealDataRowInArray = rawLastRow
        Exit Function
    End If
    If lastCol < firstCol Then
        LastRealDataRowInArray = rawLastRow
        Exit Function
    End If

    Dim maxAmount As Double
    Dim r As Long, c As Long, amt As Double, ok As Boolean
    For r = headerRow + 1 To rawLastRow
        For c = firstCol To lastCol
            amt = CoerceMoney(sourceData(r, c), ok)
            If ok And amt > maxAmount Then maxAmount = amt
        Next c
    Next r

    Dim lastRow As Long
    lastRow = rawLastRow
    Do While lastRow > headerRow
        Dim rowValues() As Variant
        ReDim rowValues(firstCol To lastCol)
        Dim rowBlank As Boolean
        rowBlank = True
        For c = firstCol To lastCol
            rowValues(c) = sourceData(lastRow, c)
            If Not (IsEmpty(rowValues(c)) Or IsError(rowValues(c)) Or _
                    Len(Trim$(CStr(rowValues(c) & ""))) = 0) Then
                rowBlank = False
            End If
        Next c

        If rowBlank Or RowValuesLookLikeSummary(rowValues, maxAmount) Then
            lastRow = lastRow - 1
        Else
            Exit Do
        End If
    Loop

    LastRealDataRowInArray = lastRow
End Function

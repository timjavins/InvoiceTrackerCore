' ONE-TIME MIGRATION. Converts existing tracker data to the intended type per column, so a
' column stops holding a mix of numbers and number-shaped text.
'
' Run once per workbook, then delete this module -- ingestion coerces new rows already
' (see CoerceValues.vb), so leaving it around only invites someone to run it again.
'
' Target types:
'   dates       -> real dates, Date format
'   amounts     -> numbers rounded to 2 decimals, Accounting format
'   everything  -> Text, which preserves leading zeros on store, invoice and req numbers
'
' A note on what "data type" means here: Excel has no per-cell type. A cell holds a value
' plus a number format, and the format is display only -- except Text, which also changes how
' Excel parses what you type next. So:
'
'   - A formula cell can safely be REFORMATTED. Giving column S the Date format leaves the
'     formula calculating; only the display changes.
'   - A formula cell must never have its VALUE rewritten. That replaces the formula with a
'     literal and it stops calculating.
'
' The script therefore formats formula columns and rewrites only data columns. Formula columns
' that return a date or an amount are listed in TenantFormulaFormats below.
'
' Nothing is written until you confirm a summary of what will change, and the workbook is
' left unsaved so an unwanted result can be discarded by closing without saving.

Public Sub MigrateColumnTypes()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(TenantSheetName("tracker"))
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Tracker sheet '" & TenantSheetName("tracker") & "' not found.", vbCritical
        Exit Sub
    End If

    ' Columns to convert, by concept so each tenant's own layout is used.
    Dim dateConcepts As Variant, moneyConcepts As Variant, textConcepts As Variant
    dateConcepts = Array("request-date", "invoice-approval-date")
    moneyConcepts = Array("total")
    textConcepts = Array("store-number", "submitted-invoice-number", "coupa-invoice-number", _
                         "requisition-number", "invoice-type", "transaction-details", _
                         "payment-number", "notes", "blocker-email", "blocker-name")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, TenantColLetter("store-number")).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data rows found on '" & ws.name & "'.", vbInformation
        Exit Sub
    End If

    ' Dry run first: count what would change without writing anything.
    Dim report As String
    Dim totalChanges As Long, totalSkipped As Long
    report = SurveyColumns(ws, lastRow, dateConcepts, "date", totalChanges, totalSkipped) & _
             SurveyColumns(ws, lastRow, moneyConcepts, "money", totalChanges, totalSkipped) & _
             SurveyColumns(ws, lastRow, textConcepts, "text", totalChanges, totalSkipped)

    If totalChanges = 0 And totalSkipped = 0 Then
        MsgBox "Nothing to convert -- every column already holds the intended type.", vbInformation
        Exit Sub
    End If

    Dim prompt As String
    prompt = "Convert existing data on '" & ws.name & "', rows 2 to " & lastRow & "?" & _
             vbCrLf & vbCrLf & report & vbCrLf & _
             "Cells to change: " & totalChanges & vbCrLf
    If totalSkipped > 0 Then
        prompt = prompt & "Cells skipped (formulas or unreadable values): " & totalSkipped & vbCrLf
    End If
    prompt = prompt & "Formula columns will be reformatted only -- their formulas stay intact." & vbCrLf
    prompt = prompt & vbCrLf & _
             "The workbook is NOT saved afterwards. Review the result, then save if it looks " & _
             "right, or close without saving to discard it."

    If MsgBox(prompt, vbYesNo + vbQuestion, "One-time type migration") <> vbYes Then Exit Sub

    ResetThinking
    On Error GoTo Failed
    PauseThinking

    ws.Activate
    UnprotectSheet

    Dim filterRow As Long
    filterRow = RemoveAutoFilter(ws)

    Dim changed As Long, skipped As Long
    changed = 0: skipped = 0

    ConvertColumns ws, lastRow, dateConcepts, "date", changed, skipped
    ConvertColumns ws, lastRow, moneyConcepts, "money", changed, skipped
    ConvertColumns ws, lastRow, textConcepts, "text", changed, skipped

    ' Formula columns: format only, never rewrite the value.
    Dim formulaFormats As Long
    formulaFormats = FormatFormulaColumns(ws, lastRow)

    If filterRow > 0 Then RestoreAutoFilter ws, filterRow
    ProtectSheet
    RestoreThinking

    Dim done As String
    done = changed & " cell(s) converted."
    If formulaFormats > 0 Then
        done = done & vbCrLf & formulaFormats & " formula column(s) reformatted " & _
               "(formulas left intact)."
    End If
    If skipped > 0 Then
        done = done & vbCrLf & skipped & " cell(s) skipped -- see the Immediate window."
    End If
    done = done & vbCrLf & vbCrLf & _
           "NEXT: re-import the Coupa exports. The tracker's identifiers are now text, and a " & _
           "lookup only matches when both sides are the same type -- ingestion writes them as " & _
           "text, so a fresh import brings the Coupa sheets into line." & vbCrLf & vbCrLf & _
           "Nothing has been saved yet."
    MsgBox done, vbInformation, "Migration complete"
    Exit Sub

Failed:
    ResetThinking
    failureState = False
    MsgBox "Migration stopped: " & Err.Description & vbCrLf & vbCrLf & _
           "Close the workbook without saving to discard any partial change.", vbCritical
End Sub

' Counts what would change, without writing. Returns a line per column for the prompt.
Private Function SurveyColumns(ByVal ws As Worksheet, ByVal lastRow As Long, _
                               ByVal concepts As Variant, ByVal kind As String, _
                               ByRef totalChanges As Long, ByRef totalSkipped As Long) As String
    Dim i As Long
    Dim col As String
    Dim out As String

    For i = LBound(concepts) To UBound(concepts)
        col = ColumnForConcept(CStr(concepts(i)))
        If Len(col) > 0 Then
            Dim changes As Long, skips As Long
            changes = 0: skips = 0
            ScanColumn ws, col, lastRow, kind, changes, skips, False

            If changes > 0 Or skips > 0 Then
                out = out & "  " & col & "  " & concepts(i) & " -> " & kind & _
                      "  (" & changes & " to change"
                If skips > 0 Then out = out & ", " & skips & " skipped"
                out = out & ")" & vbCrLf
            End If

            totalChanges = totalChanges + changes
            totalSkipped = totalSkipped + skips
        End If
    Next i

    If Len(out) > 0 Then SurveyColumns = UCase$(Left$(kind, 1)) & Mid$(kind, 2) & ":" & vbCrLf & out & vbCrLf
End Function

Private Sub ConvertColumns(ByVal ws As Worksheet, ByVal lastRow As Long, _
                           ByVal concepts As Variant, ByVal kind As String, _
                           ByRef changed As Long, ByRef skipped As Long)
    Dim i As Long
    Dim col As String
    For i = LBound(concepts) To UBound(concepts)
        col = ColumnForConcept(CStr(concepts(i)))
        If Len(col) > 0 Then
            ScanColumn ws, col, lastRow, kind, changed, skipped, True
        End If
    Next i
End Sub

' Walks one column, converting each cell to the target type. Set apply to False to count only.
'
' Works cell by cell rather than reading the column into an array, because a formula has to be
' detected per cell and the whole point is to leave those untouched.
Private Sub ScanColumn(ByVal ws As Worksheet, ByVal col As String, ByVal lastRow As Long, _
                       ByVal kind As String, ByRef changed As Long, ByRef skipped As Long, _
                       ByVal apply As Boolean)
    Dim r As Long
    Dim cell As Range
    Dim ok As Boolean

    For r = 2 To lastRow
        Set cell = ws.Cells(r, col)

        If cell.HasFormula Then
            skipped = skipped + 1
        ElseIf Len(CStr(cell.value & "")) = 0 Then
            ' Blank: set the format so future entries are typed, but count nothing.
            If apply Then cell.NumberFormat = FormatFor(kind)
        Else
            Select Case kind
                Case "date"
                    Dim d As Date
                    d = CoerceDate(cell.value, ok)
                    If ok Then
                        If Not (VarType(cell.value) = vbDate And cell.NumberFormat = FormatFor("date")) Then
                            changed = changed + 1
                            If apply Then
                                cell.NumberFormat = FormatFor("date")
                                cell.value = d
                            End If
                        End If
                    Else
                        skipped = skipped + 1
                        If apply Then
                            Debug.Print "Migrate: " & cell.Address(False, False) & " '" & _
                                        CStr(cell.value) & "' is not a date -- left as is."
                        End If
                    End If

                Case "money"
                    Dim m As Double
                    m = CoerceMoney(cell.value, ok)
                    If ok Then
                        If Not (VBA.IsNumeric(cell.value) And VarType(cell.value) <> vbString _
                                And cell.NumberFormat = FormatFor("money")) Then
                            changed = changed + 1
                            If apply Then
                                cell.NumberFormat = FormatFor("money")
                                cell.value = m
                            End If
                        End If
                    Else
                        skipped = skipped + 1
                        If apply Then
                            Debug.Print "Migrate: " & cell.Address(False, False) & " '" & _
                                        CStr(cell.value) & "' is not an amount -- left as is."
                        End If
                    End If

                Case "text"
                    Dim s As String
                    s = CoerceIdentifier(cell.value)
                    If VarType(cell.value) <> vbString Or cell.NumberFormat <> "@" _
                       Or CStr(cell.value) <> s Then
                        changed = changed + 1
                        If apply Then
                            cell.NumberFormat = "@"
                            cell.value = s
                        End If
                    End If
            End Select
        End If
    Next r
End Sub

' Applies a date or Accounting format to formula columns, leaving their formulas alone.
' Returns how many columns were reformatted.
'
' Which columns these are is per-tenant, so it comes from TenantFormulaFormats in the
' variant's TenantConfig. A tenant that has not declared any simply gets nothing done here.
Private Function FormatFormulaColumns(ByVal ws As Worksheet, ByVal lastRow As Long) As Long
    Dim spec As Variant
    On Error Resume Next
    spec = TenantFormulaFormats()
    On Error GoTo 0

    If Not IsArray(spec) Then Exit Function

    Dim i As Long
    Dim parts() As String
    For i = LBound(spec) To UBound(spec)
        parts = Split(CStr(spec(i)), "=")
        If UBound(parts) = 1 Then
            Dim col As String, kind As String
            col = Trim$(parts(0))
            kind = LCase$(Trim$(parts(1)))

            On Error Resume Next
            ws.Range(col & "2:" & col & lastRow).NumberFormat = FormatFor(kind)
            If Err.Number = 0 Then FormatFormulaColumns = FormatFormulaColumns + 1
            Err.Clear
            On Error GoTo 0
        End If
    Next i
End Function

' Accounting format, matching what Excel's Accounting preset applies for USD.
Private Function FormatFor(ByVal kind As String) As String
    Select Case kind
        Case "date":  FormatFor = "mm/dd/yyyy"
        Case "money": FormatFor = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        Case Else:    FormatFor = "@"
    End Select
End Function

' Resolves a concept to a column letter, returning "" when this tenant has no such column.
Private Function ColumnForConcept(ByVal concept As String) As String
    On Error Resume Next
    ColumnForConcept = TenantColLetter(concept)
    On Error GoTo 0
End Function

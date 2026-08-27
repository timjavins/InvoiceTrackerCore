' Turns supplier store values that are not plain numbers -- "0391-A", "2242-B" -- into real
' 4-digit Nordstrom store numbers by asking the user, once per distinct value, before an
' import writes any rows.
'
' Why ask rather than derive: the suffix is not noise. The July monitoring file listed "0391"
' and "0391-A" as separate lines, so "-A" is a real second location or panel at store 0391.
' Nothing in the data says which store a suffix belongs to, and the arithmetic guess this
' replaces -- Right$("0000" & text, 4) -- turned "0391-A" into "91-A".
'
' Why up front rather than per row: NormalizeStoreNumber is a pure per-cell function called
' once per row, so prompting inside it would fire once per occurrence, and it runs under
' PauseThinking where a modal form paints wrong.
'
' Why cancel aborts the whole import: the answers are collected before any row is written, so
' there is nothing to roll back. Continuing with a half-answered map would write unresolved
' values into the tracker, which is the state this whole file exists to prevent.
'
' Answers are deliberately NOT persisted. A suffix can mean something different next month,
' and a remembered guess nobody re-reads is how "91-A" survived in the tracker for as long as
' it did.

' Builds raw value -> 4-digit store number for every non-conforming store on the source sheet.
'
' Returns False only when the user cancelled, which the caller must treat as abort. An empty
' resolutions map with a True result is the normal case: almost every file has nothing odd.
Public Function BuildStoreResolutions(ByVal wsSource As Worksheet, _
                                      ByVal headerRow As Long, _
                                      ByVal lastRow As Long, _
                                      ByVal sourceColIndexes As Object, _
                                      ByRef resolutions As Object) As Boolean
    Set resolutions = CreateObject("Scripting.Dictionary")

    ' Nothing to inspect is not a failure -- let the import proceed and fail on its own terms.
    BuildStoreResolutions = True
    If wsSource Is Nothing Then Exit Function
    If sourceColIndexes Is Nothing Then Exit Function
    If Not sourceColIndexes.Exists("STORE #") Then Exit Function

    Dim pending As Object
    Set pending = CollectNonConformingStores(wsSource, headerRow, lastRow, _
                                             CLng(sourceColIndexes("STORE #")))
    If pending.Count = 0 Then Exit Function

    ' Columns worth showing the user. None are required by the import -- the tracker gets
    ' address fields from 'Store Directory' -- so they are looked up by name and skipped when
    ' this supplier did not include them.
    Dim ctxCols As Object
    Set ctxCols = GetHeaderColumnIndexes(wsSource, headerRow, _
        Array("STORE NAME", "STORE ADDRESS", "CITY", "ST", "ZIP", _
              "TRANSACTION DETAILS", "COMMENTS", "NOTES", _
              "TOTAL", "SUBTOTAL", "INV DATE", "INVOICE #", "INVOICE TYPE", "BILL CODE"))

    ' The dialog is modal, so screen updating has to be back on for it to paint.
    '
    ' Conditional, not a bare RestoreThinking: if no pause is active there is no saved state to
    ' restore and the unwind would try to set Calculation = 0. AddNewBills always pauses first, but
    ' this must not depend on its only caller doing so.
    Dim wasPaused As Boolean
    wasPaused = SuspendThinkingForDialog()

    Dim rawValue As Variant
    Dim answer As String
    For Each rawValue In pending.Keys
        answer = PromptStoreNumberFix(StoreFixContext(wsSource, CLng(pending(rawValue)), _
                                                      CStr(rawValue), ctxCols))

        ' Empty means cancelled or closed. Belt and braces on the format: the form validates
        ' input, but core must not write whatever a form happened to hand back.
        answer = NormalizedStoreEntry(answer)
        If Len(answer) = 0 Then
            ResumeThinkingAfterDialog wasPaused
            BuildStoreResolutions = False
            Exit Function
        End If

        resolutions(CStr(rawValue)) = answer
    Next rawValue

    ResumeThinkingAfterDialog wasPaused
End Function

' Accepts 1 to 4 digits and nothing else, returning the value zero-padded to the 4-digit
' convention in CONTEXT.md. Returns "" for anything it will not vouch for.
Public Function NormalizedStoreEntry(ByVal entry As String) As String
    Dim text As String
    text = Trim$(entry)

    If Len(text) = 0 Then Exit Function
    If Len(text) > 4 Then Exit Function

    Dim i As Long
    Dim ch As String
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i

    NormalizedStoreEntry = Right$("0000" & text, 4)
End Function

' Packs the fields a human needs in order to say what a funky store value means.
'
' A Dictionary rather than a Type because the stacked code is pasted into ThisWorkbook, a
' class module, where Public Type is not allowed.
Private Function StoreFixContext(ByVal wsSource As Worksheet, _
                                 ByVal rowIndex As Long, _
                                 ByVal rawValue As String, _
                                 ByVal ctxCols As Object) As Object
    Dim ctx As Object
    Set ctx = CreateObject("Scripting.Dictionary")
    Set StoreFixContext = ctx

    ctx("RawValue") = rawValue
    ctx("SheetName") = wsSource.Name
    ctx("RowIndex") = rowIndex
    ctx("Address") = StoreFixAddress(wsSource, rowIndex, ctxCols)
    ctx("Details") = JoinNonEmpty(vbCrLf, _
        ContextText(wsSource, rowIndex, ctxCols, "TRANSACTION DETAILS"), _
        ContextText(wsSource, rowIndex, ctxCols, "COMMENTS"), _
        ContextText(wsSource, rowIndex, ctxCols, "NOTES"))
    ctx("Amount") = ContextText(wsSource, rowIndex, ctxCols, "TOTAL")
    If Len(ctx("Amount")) = 0 Then
        ctx("Amount") = ContextText(wsSource, rowIndex, ctxCols, "SUBTOTAL")
    End If
    ctx("InvDate") = ContextText(wsSource, rowIndex, ctxCols, "INV DATE")
    ctx("InvoiceNumber") = JoinNonEmpty(" / ", _
        ContextText(wsSource, rowIndex, ctxCols, "INVOICE #"), _
        ContextText(wsSource, rowIndex, ctxCols, "BILL CODE"))
    ctx("InvoiceType") = ContextText(wsSource, rowIndex, ctxCols, "INVOICE TYPE")
End Function

' Assembles whatever address columns this supplier happened to include, on one line.
Private Function StoreFixAddress(ByVal ws As Worksheet, _
                                 ByVal rowIndex As Long, _
                                 ByVal ctxCols As Object) As String
    Dim streetPart As String
    streetPart = JoinNonEmpty(", ", _
        ContextText(ws, rowIndex, ctxCols, "STORE NAME"), _
        ContextText(ws, rowIndex, ctxCols, "STORE ADDRESS"), _
        ContextText(ws, rowIndex, ctxCols, "CITY"))

    ' State and ZIP read as "WA 98101", not "WA, 98101".
    Dim regionPart As String
    regionPart = JoinNonEmpty(" ", _
        ContextText(ws, rowIndex, ctxCols, "ST"), _
        ContextText(ws, rowIndex, ctxCols, "ZIP"))

    StoreFixAddress = JoinNonEmpty(", ", streetPart, regionPart)
End Function

' Cell text for an optional context column, or "" when this supplier omitted it.
Private Function ContextText(ByVal ws As Worksheet, _
                             ByVal rowIndex As Long, _
                             ByVal ctxCols As Object, _
                             ByVal headerName As String) As String
    If ctxCols Is Nothing Then Exit Function
    If Not ctxCols.Exists(headerName) Then Exit Function

    ' .Text rather than .Value so dates and money arrive formatted the way the supplier
    ' displays them, which is what the user is reading off the invoice.
    ContextText = Trim$(CStr(ws.Cells(rowIndex, CLng(ctxCols(headerName))).Text))
End Function

' Joins only the parts that have content, so absent columns leave no stray separators.
Private Function JoinNonEmpty(ByVal separator As String, ParamArray parts() As Variant) As String
    Dim out As String
    Dim i As Long
    Dim part As String

    For i = LBound(parts) To UBound(parts)
        part = Trim$(CStr(parts(i)))
        If Len(part) > 0 Then
            If Len(out) > 0 Then out = out & separator
            out = out & part
        End If
    Next i

    JoinNonEmpty = out
End Function

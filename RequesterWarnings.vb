' Data-quality warnings raised while resolving requesters, and the report they are written to.
'
' WHY A FILE INSTEAD OF A MESSAGE BOX
'
' Both generators used to concatenate every warning into one "Error Summary" MsgBox. That works at
' two or three warnings. At fourteen it is a wall of text with no scrollbar and no way to sort,
' filter, or copy a single line out; past a few dozen VBA truncates the string outright and the
' warnings simply vanish. Every warning of this kind has the same shape -- a tracker row, a store, a
' bill, and what went wrong -- so it belongs in a table, not in a dialog. The dialog now carries one
' line saying how many there were and where to find them.
'
' WHY CORE DECIDES WHAT COUNTS AS A WARNING
'
' RecordRequesterOutcome takes the raw outcome of the cascade and works out both whether it is worth
' reporting at all and how to label it. The generators report facts; the judgement lives here, next
' to the cascade in RequesterEmail.vb that produced it. Otherwise each tenant would classify for
' itself, the two would drift, and a third tenant would have to reinvent the same rules -- which is
' precisely the drift ADR-0005 was written to end.
'
' A fully clean resolution -- the site's own responsible party, nothing skipped -- is NOT a warning
' and is not recorded. A report that also listed every row that worked would be as unreadable as
' the dialog it replaces.
'
' THE ISSUE COLUMN IS A CLOSED SET
'
'     NoSiteRPsRow            the store is not on 'Site RPs' at all, so the bill was skipped
'     NoContactAtAnyTier      the store is listed but carries no email in any of the four columns
'     AllTiersInactive        addresses exist, but not one of them is an active Coupa user
'     EscalatedPastInactive   exported fine, but not under the site's own responsible party
'     StoreNotInBUList        the store has no GL code, so the bill was skipped
'
' Five stable values with no row-specific text in them, so the user can filter the column and see
' all of one kind at once. Everything row-specific goes in Detail.
'
' CLASSIFIED FROM A CODE, NEVER FROM THE PROSE
'
' TryResolveSiteRequester returns both a human sentence (reason) and a stable machine code
' (outcome). Only the code is matched on here. Pattern-matching the sentence would make its wording
' load-bearing: reword a warning for clarity and the classification silently changes, with no
' compiler and no test to catch it.

' The warnings recorded so far, one Variant array per warning in WarningColumnHeaders order.
'
' A Collection rather than a Dictionary because there is no key: two bills on the same store can
' both be worth reporting, and the natural order is the order the generator met them, which is
' tracker order.
Private requesterWarnings As Collection

' Clears the warning list. Call at the start of a run, before the row loop.
'
' Reset explicitly rather than lazily on first record, because the count and the file both have to
' describe THIS run. A second export in the same session must not inherit the first one's warnings.
Public Sub ResetRequesterWarnings()
    Set requesterWarnings = New Collection
End Sub

' Records the outcome of a requester lookup, if it is worth recording.
'
' outcome is the code from TryResolveSiteRequester. requesterEmail is the address the row actually
' went out under, which on a no-contact store is the prompted default -- what the user needs to know
' is who the requisition landed on, not that the field was empty mid-lookup.
'
' Silently returns for a clean resolution. Callers therefore do not have to know which outcomes
' matter; they hand over what happened and let this decide.
Public Sub RecordRequesterOutcome(ByVal trackerRow As Long, _
                                  ByVal store As String, _
                                  ByVal billRef As String, _
                                  ByVal outcome As String, _
                                  ByVal requesterEmail As String, _
                                  ByVal tier As String, _
                                  ByVal inactiveSkipped As String, _
                                  ByVal detail As String)
    Dim issue As String
    issue = IssueForOutcome(outcome)
    If Len(issue) = 0 Then Exit Sub

    AppendRequesterWarning trackerRow, store, billRef, issue, _
                           requesterEmail, tier, inactiveSkipped, detail
End Sub

' Records a bill skipped because its store has no row on 'BU List'.
'
' Separate from RecordRequesterOutcome because it is not an outcome of the cascade -- the store fails
' the GL lookup before a requester is ever asked for. It lands in the same report because it is the
' same kind of problem with the same fix: one row in a shared list. Securitas raises it from three
' places; JCI's generator does not consult 'BU List' at all and never will.
Public Sub RecordStoreNotInBUList(ByVal trackerRow As Long, _
                                  ByVal store As String, _
                                  ByVal billRef As String, _
                                  ByVal detail As String)
    AppendRequesterWarning trackerRow, store, billRef, "StoreNotInBUList", _
                           vbNullString, vbNullString, vbNullString, detail
End Sub

' How many warnings this run produced. 0 means the run was clean.
Public Function RequesterWarningCount() As Long
    If requesterWarnings Is Nothing Then Exit Function
    RequesterWarningCount = requesterWarnings.Count
End Function

' Writes the warnings to an .xlsx in Downloads and returns the path, or "" when nothing was written.
'
' Returns "" for a CLEAN RUN and writes no file. This is deliberate and load-bearing: an empty
' report, or worse a file left over from last week, would be read as this run's verdict. No file
' means no warnings, and the caller says so.
'
' Also returns "" when the destination is locked, after saying which file to close -- the same
' behaviour, and the same wording, the generators already use for a Coupa upload CSV that is still
' open in Excel. Warnings are then reported by count alone rather than lost.
'
' Lands beside the run's other artifact ('Coupa uploads YYYYMMDD.csv', same folder, same date shape)
' so both outputs of one export sit together in Downloads.
Public Function WriteRequesterWarningsReport() As String
    If RequesterWarningCount() = 0 Then Exit Function

    Dim outputPath As String
    outputPath = RequesterWarningsPath()

    If IsFileWriteLocked(outputPath) Then
        ' The generators run paused, so a modal dialog would not paint.
        Dim wasPaused As Boolean
        wasPaused = SuspendThinkingForDialog()
        MsgBox "Your previous requester warnings file '" & outputPath & "' is open in Excel or " & _
               "another program, so this run's warnings could not be written. Please close it " & _
               "and try again.", vbExclamation, "File Creation Error"
        ResumeThinkingAfterDialog wasPaused
        Exit Function
    End If

    Dim headers As Variant
    headers = WarningColumnHeaders()

    Dim colCount As Long
    colCount = UBound(headers) - LBound(headers) + 1

    ' Header row plus one row per warning, filled in memory and written in a single assignment.
    ' A per-cell loop over eight columns is eight worksheet writes per warning for no gain.
    Dim grid() As Variant
    ReDim grid(1 To requesterWarnings.Count + 1, 1 To colCount)

    Dim c As Long
    For c = 1 To colCount
        grid(1, c) = headers(LBound(headers) + c - 1)
    Next c

    Dim r As Long
    Dim entry As Variant
    r = 1
    For Each entry In requesterWarnings
        r = r + 1
        For c = 1 To colCount
            grid(r, c) = entry(c - 1)
        Next c
    Next entry

    Dim wbReport As Workbook
    Dim wsReport As Worksheet

    ' xlWBATWorksheet forces exactly one sheet regardless of the user's SheetsInNewWorkbook setting,
    ' so the report never ships with two empty sheets after it.
    Set wbReport = Workbooks.Add(xlWBATWorksheet)

    On Error GoTo WriteFailed

    Set wsReport = wbReport.Sheets(1)
    wsReport.Name = "Requester warnings"

    ' Text format BEFORE the values land. Store numbers and bill codes are identifiers, not
    ' quantities: written into a General cell, "0515" becomes 515 and stops matching the tracker
    ' the user is comparing against.
    wsReport.Columns(2).NumberFormat = "@"
    wsReport.Columns(3).NumberFormat = "@"

    wsReport.Range(wsReport.Cells(1, 1), _
                   wsReport.Cells(requesterWarnings.Count + 1, colCount)).Value = grid

    wsReport.Rows(1).Font.Bold = True
    wsReport.Rows(1).Interior.Color = RGB(217, 217, 217)

    ' Freeze through the window rather than by selecting row 2 and setting ActiveWindow.FreezePanes.
    ' Selecting requires the sheet to be active, which a generator running with screen updating off
    ' cannot rely on -- and it would move the user's selection in whatever workbook they were in.
    With wbReport.Windows(1)
        .SplitRow = 1
        .SplitColumn = 0
        .FreezePanes = True
    End With

    ' Filter across the whole used range, so filtering Issue hides the other rows entirely rather
    ' than leaving their trailing columns visible.
    wsReport.Range(wsReport.Cells(1, 1), _
                   wsReport.Cells(requesterWarnings.Count + 1, colCount)).AutoFilter

    wsReport.Columns(1).Resize(, colCount).AutoFit

    ' Detail and Inactive Skipped hold sentences and address lists, and AutoFit on a 200-character
    ' sentence produces a column wider than the screen -- which is the unreadable-wall-of-text
    ' problem this report exists to solve, moved into a spreadsheet. Cap those two; the rest are
    ' short enough to leave alone.
    For c = 1 To colCount
        If c >= colCount - 1 Then
            If wsReport.Columns(c).ColumnWidth > 60 Then wsReport.Columns(c).ColumnWidth = 60
        End If
    Next c

    ' SaveAs over an existing file prompts unless alerts are off, and a prompt behind a paused Excel
    ' session is an invisible hang. Restore whatever the caller had rather than assuming False.
    Dim priorAlerts As Boolean
    priorAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wbReport.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = priorAlerts

    wbReport.Close SaveChanges:=False

    WriteRequesterWarningsReport = outputPath
    Exit Function

WriteFailed:
    ' Never leave an orphan workbook on screen: the generator that called this is mid-run with its
    ' own cleanup to do, and an unsaved book in the way turns one failure into a confusing second
    ' one. The warnings stay in the list, so the caller can still report the count.
    Dim failureDetail As String
    failureDetail = Err.Description

    On Error Resume Next
    Application.DisplayAlerts = False
    wbReport.Close SaveChanges:=False
    On Error GoTo 0

    Dim dialogPaused As Boolean
    dialogPaused = SuspendThinkingForDialog()
    MsgBox "The requester warnings report could not be written: " & failureDetail & vbCrLf & _
           vbCrLf & "The export itself is unaffected.", vbExclamation, "Requester Warnings"
    ResumeThinkingAfterDialog dialogPaused
End Function

' Adds one warning to the list, in WarningColumnHeaders order.
'
' Tracker Row stays a Long so the report sorts numerically; everything else is text.
Private Sub AppendRequesterWarning(ByVal trackerRow As Long, _
                                   ByVal store As String, _
                                   ByVal billRef As String, _
                                   ByVal issue As String, _
                                   ByVal requesterEmail As String, _
                                   ByVal tier As String, _
                                   ByVal inactiveSkipped As String, _
                                   ByVal detail As String)
    If requesterWarnings Is Nothing Then ResetRequesterWarnings

    requesterWarnings.Add Array(trackerRow, store, billRef, issue, _
                                requesterEmail, tier, inactiveSkipped, detail)
End Sub

' The Issue value for a cascade outcome code, or "" when the outcome is not worth reporting.
'
' "resolved-primary" is the whole reason this returns "": the site's own responsible party answered
' and is an active Coupa user, which is the normal case and not news.
'
' An unrecognized code is reported under its own raw value rather than dropped. Only core emits
' these, so an unknown one means core changed on one side of this mapping and not the other -- a bug
' that should be visible in the report rather than swallowed there.
Private Function IssueForOutcome(ByVal outcome As String) As String
    Select Case LCase$(Trim$(outcome))
        Case "store-missing":      IssueForOutcome = "NoSiteRPsRow"
        Case "no-contacts":        IssueForOutcome = "NoContactAtAnyTier"
        Case "all-inactive":       IssueForOutcome = "AllTiersInactive"
        Case "resolved-escalated": IssueForOutcome = "EscalatedPastInactive"
        Case "resolved-primary":   IssueForOutcome = vbNullString
        Case Else:                 IssueForOutcome = Trim$(outcome)
    End Select
End Function

' Where the report is written.
'
' Downloads and the YYYYMMDD suffix match 'Coupa uploads YYYYMMDD.csv' on purpose, so a run's two
' artifacts sort next to each other and the date is readable without opening either.
Private Function RequesterWarningsPath() As String
    RequesterWarningsPath = Environ("USERPROFILE") & "\Downloads\Coupa requester warnings " & _
                            Format(Date, "YYYYMMDD") & ".xlsx"
End Function

' The report's columns, in order. The array's order IS the schema -- AppendRequesterWarning and
' WriteRequesterWarningsReport both index against it.
'
' Named "...Headers" and not "warningColumns" for the reason spelled out at TierLabelNames in
' RequesterEmail.vb: a procedure whose name matches some caller's local variable is the same
' identifier to VBA and fails at run time, not at compile time.
Private Function WarningColumnHeaders() As Variant
    WarningColumnHeaders = Array("Tracker Row", "Store", "Bill Ref", "Issue", _
                                 "Requester Used", "Tier", "Inactive Skipped", "Detail")
End Function

' Whether a file exists and is held open by something else.
'
' Same probe the generators use before saving a Coupa upload CSV: opening for binary read-write with
' a lock succeeds only if nothing else holds the file. Tested up front because SaveAs on a locked
' path raises a generic automation error that says nothing about which file or why.
Private Function IsFileWriteLocked(ByVal filePath As String) As Boolean
    ' FileSystemObject rather than Dir(), which keeps global iteration state a caller mid-Dir-loop
    ' would lose.
    Dim fsoCheck As Object
    Set fsoCheck = CreateObject("Scripting.FileSystemObject")
    If Not fsoCheck.FileExists(filePath) Then Exit Function

    Dim testFile As Integer
    testFile = FreeFile()

    On Error Resume Next
    Open filePath For Binary Access Read Write Lock Read Write As #testFile
    If Err.Number <> 0 Then IsFileWriteLocked = True
    Err.Clear
    Close #testFile
    On Error GoTo 0
End Function

' Which bill files this workbook has fully processed, keyed by content hash (HashFile).
'
' The log is a hidden sheet in the WORKING tracker workbook itself, not a file beside it:
' the tracker is what gets shared, copied and rolled over each year, and the log has to
' travel with it. The sheet is created on first write, so rolling this out needs no manual
' migration step. Sheet name comes from TenantProcessedBatchSheet(). See the batch-tracking
' spec, section 1.
'
' Row layout: HASH | SOURCE FILE | TENANT | PROCESSED AT | ROWS ADDED | ROWS SKIPPED

Private Const PBL_HEADERS As String = "HASH|SOURCE FILE|TENANT|PROCESSED AT|ROWS ADDED|ROWS SKIPPED"

Public Function IsAlreadyProcessed(ByVal wb As Workbook, ByVal hash As String) As Boolean
    IsAlreadyProcessed = (PblFindRow(wb, hash) > 0)
End Function

' One line describing the earlier run of this file, for the "already processed" message.
Public Function ProcessedBatchSummary(ByVal wb As Workbook, ByVal hash As String) As String
    Dim r As Long
    r = PblFindRow(wb, hash)
    If r = 0 Then Exit Function

    With wb.Worksheets(TenantProcessedBatchSheet())
        ProcessedBatchSummary = CStr(.Cells(r, 2).Value) & ", processed " & _
            Format$(.Cells(r, 4).Value, "mm/dd/yyyy h:nn AM/PM") & ": " & _
            CStr(.Cells(r, 5).Value) & " added, " & CStr(.Cells(r, 6).Value) & " skipped"
    End With
End Function

Public Sub RecordProcessedBatch(ByVal wb As Workbook, ByVal hash As String, _
                                ByVal sourceFile As String, ByVal tenant As String, _
                                ByVal rowsAdded As Long, ByVal rowsSkipped As Long)
    Dim ws As Worksheet
    Dim r As Long

    Set ws = PblEnsureSheet(wb)
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(r, 1).NumberFormat = "@"
    ws.Cells(r, 1).Value = LCase$(Trim$(hash))
    ws.Cells(r, 2).Value = sourceFile
    ws.Cells(r, 3).Value = tenant
    ws.Cells(r, 4).Value = Now
    ws.Cells(r, 5).Value = rowsAdded
    ws.Cells(r, 6).Value = rowsSkipped
End Sub

Private Function PblFindRow(ByVal wb As Workbook, ByVal hash As String) As Long
    Dim key As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    key = LCase$(Trim$(hash))
    If Len(key) = 0 Then Exit Function
    If Not PblSheetExists(wb, TenantProcessedBatchSheet()) Then Exit Function

    Set ws = wb.Worksheets(TenantProcessedBatchSheet())
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If LCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = key Then
            PblFindRow = r
            Exit Function
        End If
    Next r
End Function

Private Function PblEnsureSheet(ByVal wb As Workbook) As Worksheet
    Dim name As String
    Dim ws As Worksheet
    Dim prior As Object
    Dim headers As Variant

    name = TenantProcessedBatchSheet()
    If PblSheetExists(wb, name) Then
        Set PblEnsureSheet = wb.Worksheets(name)
        Exit Function
    End If

    ' Adding a sheet activates it; put the user back where they were.
    Set prior = wb.ActiveSheet
    Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = name
    headers = Split(PBL_HEADERS, "|")
    ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(headers) + 1)).Value = headers
    ws.Visible = xlSheetHidden

    On Error Resume Next
    If Not prior Is Nothing Then prior.Activate
    On Error GoTo 0

    Set PblEnsureSheet = ws
End Function

Private Function PblSheetExists(ByVal wb As Workbook, ByVal name As String) As Boolean
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, name, vbTextCompare) = 0 Then
            PblSheetExists = True
            Exit Function
        End If
    Next ws
End Function

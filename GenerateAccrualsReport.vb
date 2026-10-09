' This subroutine generates an accruals report by filtering records from the tracker sheet
' based on specific criteria and exporting matching records to a new Excel workbook.
'
' Dormant unless the tenant opts in: TenantAccrualsEnabled() must return True. A tracker that has not
' agreed its accrual criteria with finance gets a notice instead of a report built on guesses.
'
' Selection Criteria:
' 1. INVOICE TYPE is not one of TenantAccrualExcludedInvoiceTypes()
' 2. INV Approval Date is blank or after the cut-off date (the last day of the fiscal month chosen
'    at the prompt; see AskForCutoffDate)
' 3. BU is one of TenantAccrualBusinessUnits()
' 4. REQ # is blank or numeric (text values like "N/A" or "PENDING" are excluded)
' 5. PO STATUS does not contain "Cancelled" or "Closed"
' 6. ORDER DATE is blank or after (cut-off date - 13 weeks) [13-week sheet only]
'
' Output: New workbook with two sheets:
'   "T - 3 mos." - records matching all 6 criteria (13-week window)
'   "YTD"        - records matching criteria 1-5 (full fiscal year to cutoff)
' Workbook saved to "Downloads" folder.
'
' Columns are read through TenantColLetter concepts, never by letter. Rows whose formula columns hold
' #N/A are excluded rather than crashing the run -- see the comment on the filter loop for which way
' each criterion resolves an error value.

Sub GenerateAccrualsReport()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cutoffDate As Date
    Dim thresholdDate As Date
    Dim i As Long
    Dim j As Long
    Dim h As Long
    Dim matchCount As Long

    ' One source array per output column, in output order. Each entry is the 2D array ColumnValues
    ' returns for that concept's tracker column.
    Dim concepts As Variant
    concepts = Array("store-number", "business-unit", "invoice-type", "submitted-invoice-number", _
                     "total", "requisition-number", "requisition-status", "purchase-order-number", _
                     "order-date", "purchase-order-status", "invoice-status", "invoice-approval-date", "notes")
    Dim headers As Variant
    headers = Array("STORE #", "BU", "INVOICE TYPE", "BILL CODE", "TOTAL", "REQ #", "REQ STATUS", "PO #", _
                    "ORDER DATE", "PO STATUS", "INV STATUS", "INV Approval Date", "NOTES/COMMENTS")
    Const COL_BU As Long = 2
    Const COL_TYPE As Long = 3
    Const COL_REQ As Long = 6
    Const COL_ORDER_DATE As Long = 9
    Const COL_PO_STATUS As Long = 10
    Const COL_APPROVAL As Long = 12
    Dim src(1 To 13) As Variant

    ' 13-week output array
    Dim outputArr() As Variant
    Dim outputRow As Long

    ' YTD output array
    Dim outputArrYTD() As Variant
    Dim outputRowYTD As Long
    Dim matchCountYTD As Long

    ' New workbook variables
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim wsYTD As Worksheet
    Dim savePath As String
    Dim fileName As String

    ' Loop helpers
    Dim reqVal As String
    Dim orderDate As Date
    Dim orderDateBlank As Boolean
    Dim approvalText As String
    Dim buText As String
    Dim typeText As String
    Dim poStatus As String
    Dim orderText As String
    Dim excludedTypes As Variant
    Dim businessUnits As Variant
    Dim excluded As Boolean
    Dim inScope As Boolean
    Dim k As Long

    If Not TenantAccrualsEnabled() Then
        MsgBox "The accruals report is not set up for this tracker.", vbInformation, "Accruals Report"
        Exit Sub
    End If

    ' Clear anything a previous run left behind, so the pause below is the outermost one.
    ResetThinking
    ResetProtection

    ' Calculation and screen updating must come back on however this sub ends.
    '
    ' This used to set the two Application properties directly with no error handler, so any failure
    ' below left the workbook on manual calculation with nothing to explain why. Armed this early so a
    ' tenant config error (an unknown column concept) is reported the same way as any other failure.
    On Error GoTo Failed

    Set ws = ThisWorkbook.Sheets(TenantSheetName("tracker"))
    lastRow = ws.Cells(ws.Rows.Count, TenantCol("store-number")).End(xlUp).Row

    ' Nothing below the header row. Every ReDim and UBound below assumes at least one data row, and
    ' a single data row would make Range.Value return a scalar rather than an array.
    If lastRow < 2 Then
        MsgBox "There are no invoice rows to report on.", vbInformation, "Accruals Report"
        Exit Sub
    End If

    ' The cut-off is the last day of a fiscal month the user picks from a list.
    cutoffDate = AskForCutoffDate()
    If CDbl(cutoffDate) = 0 Then
        MsgBox "Report generation cancelled.", vbInformation
        Exit Sub
    End If
    thresholdDate = DateAdd("ww", -13, cutoffDate)

    Debug.Print "Cut-off Date: " & cutoffDate
    Debug.Print "13-Week Threshold: " & thresholdDate

    excludedTypes = TenantAccrualExcludedInvoiceTypes()
    businessUnits = TenantAccrualBusinessUnits()

    PauseThinking

    ' Read all relevant columns into arrays. ColumnValues guarantees a 2D array even when the
    ' tracker has exactly one data row.
    For h = 1 To 13
        src(h) = ColumnValues(ws, TenantColLetter(CStr(concepts(h - 1))), 2, lastRow)
    Next h

    ' Initialize output arrays (max size = source data size)
    ReDim outputArr(1 To UBound(src(1), 1) + 1, 1 To 13)
    ReDim outputArrYTD(1 To UBound(src(1), 1) + 1, 1 To 13)

    ' Set headers in first row of both arrays
    For h = 1 To 13
        outputArr(1, h) = headers(h - 1)
        outputArrYTD(1, h) = headers(h - 1)
    Next h

    outputRow = 2
    matchCount = 0
    outputRowYTD = 2
    matchCountYTD = 0

    ' Filter records based on criteria.
    '
    ' Every comparison below goes through SafeText. Business unit, order date and PO status are
    ' formula columns that hold #N/A whenever a lookup misses -- a store number absent from 'BU List'
    ' is enough -- and a Variant holding an error subtype raises a type mismatch on any comparison,
    ' including `colB(i, 1) = 200`. That was an unhandled crash that also left calculation on manual.
    '
    ' An error value is treated as blank rather than as a match. For BU that means the row is
    ' excluded, which is correct: a row whose business unit cannot be determined must not be
    ' reported as an in-scope BU. For the approval date it means the row is treated as not yet
    ' approved and so is accrued, which is the conservative accounting choice.
    For i = 1 To UBound(src(1), 1)

        ' Criterion 1: INVOICE TYPE is not an excluded type
        typeText = UCase$(SafeText(src(COL_TYPE)(i, 1)))
        excluded = False
        For k = LBound(excludedTypes) To UBound(excludedTypes)
            If typeText = UCase$(CStr(excludedTypes(k))) Then excluded = True
        Next k
        If excluded Then GoTo NextRow

        ' Criterion 2: INV Approval Date is blank or after cut-off date
        approvalText = SafeText(src(COL_APPROVAL)(i, 1))
        If Len(approvalText) > 0 Then
            If Not IsDate(approvalText) Then GoTo NextRow
            If CDate(approvalText) <= cutoffDate Then GoTo NextRow
        End If

        ' Criterion 3: BU is one of the tenant's in-scope business units
        buText = SafeText(src(COL_BU)(i, 1))
        inScope = False
        For k = LBound(businessUnits) To UBound(businessUnits)
            If buText = CStr(businessUnits(k)) Then inScope = True
        Next k
        If Not inScope Then GoTo NextRow

        ' Criterion 4: REQ # is blank or numeric (skip non-blank text values)
        reqVal = SafeText(src(COL_REQ)(i, 1))
        If reqVal <> "" And Not IsNumeric(reqVal) Then
            GoTo NextRow
        End If

        ' Criterion 5: PO STATUS does not contain "Cancelled" or "Closed"
        poStatus = SafeText(src(COL_PO_STATUS)(i, 1))
        If InStr(1, poStatus, "Cancelled", vbTextCompare) > 0 Or _
           InStr(1, poStatus, "Closed", vbTextCompare) > 0 Then
            GoTo NextRow
        End If

        ' Criterion 6: validate ORDER DATE; skip both sheets if the value is not a date.
        ' IsDate rather than an On Error Resume Next around CDate, so the loop never disarms the
        ' handler armed above.
        orderText = SafeText(src(COL_ORDER_DATE)(i, 1))
        orderDateBlank = (Len(orderText) = 0)
        If Not orderDateBlank Then
            If Not IsDate(orderText) Then
                Debug.Print "Row " & (i + 1) & ": ORDER DATE is not a date (" & orderText & ")"
                GoTo NextRow
            End If
            orderDate = CDate(orderText)
        End If

        ' Criteria 1-5 passed and ORDER DATE valid — add to YTD (no date window for YTD)
        For h = 1 To 13
            outputArrYTD(outputRowYTD, h) = src(h)(i, 1)
        Next h
        outputRowYTD = outputRowYTD + 1
        matchCountYTD = matchCountYTD + 1

        ' Also add to 13-week if ORDER DATE is blank or within the 13-week window
        If orderDateBlank Or orderDate > thresholdDate Then
            For h = 1 To 13
                outputArr(outputRow, h) = src(h)(i, 1)
            Next h
            outputRow = outputRow + 1
            matchCount = matchCount + 1
        End If

NextRow:
    Next i

    ' Check if any records matched at all
    If matchCount = 0 And matchCountYTD = 0 Then
        RestoreThinking
        MsgBox "No records match the specified criteria.", vbInformation, "Accruals Report"
        Exit Sub
    End If

    ' Build properly-sized final arrays (header row + data rows)
    Dim finalArr() As Variant
    ReDim finalArr(1 To outputRow - 1, 1 To 13)
    For i = 1 To outputRow - 1
        For j = 1 To 13
            finalArr(i, j) = outputArr(i, j)
        Next j
    Next i

    Dim finalArrYTD() As Variant
    ReDim finalArrYTD(1 To outputRowYTD - 1, 1 To 13)
    For i = 1 To outputRowYTD - 1
        For j = 1 To 13
            finalArrYTD(i, j) = outputArrYTD(i, j)
        Next j
    Next i

    ' Create new workbook with two sheets
    Set wbNew = Workbooks.Add
    Set wsNew = wbNew.Sheets(1)
    wsNew.Name = "T - 3 mos."
    Set wsYTD = wbNew.Sheets.Add(After:=wsNew)
    wsYTD.Name = "YTD"

    ' Write data to both sheets
    wsNew.Range(wsNew.Cells(1, 1), wsNew.Cells(outputRow - 1, 13)).Value = finalArr
    wsYTD.Range(wsYTD.Cells(1, 1), wsYTD.Cells(outputRowYTD - 1, 13)).Value = finalArrYTD

    ' Format 13-week sheet
    wsNew.Activate
    With wsNew
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(217, 217, 217)
        .Columns(1).NumberFormat = "0000"
        .Columns(5).NumberFormat = "$#,##0.00"
        .Columns(9).NumberFormat = "mm/dd/yyyy"
        .Columns(12).NumberFormat = "mm/dd/yyyy"
        .Columns("A:M").AutoFit
        .Rows(2).Select
        ActiveWindow.FreezePanes = True
        .Cells(1, 1).Select
        .Range("A1:M1").AutoFilter
    End With

    ' Format YTD sheet
    wsYTD.Activate
    With wsYTD
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(217, 217, 217)
        .Columns(1).NumberFormat = "0000"
        .Columns(5).NumberFormat = "$#,##0.00"
        .Columns(9).NumberFormat = "mm/dd/yyyy"
        .Columns(12).NumberFormat = "mm/dd/yyyy"
        .Columns("A:M").AutoFit
        .Rows(2).Select
        ActiveWindow.FreezePanes = True
        .Cells(1, 1).Select
        .Range("A1:M1").AutoFilter
    End With

    ' Leave user on the 13-week sheet
    wsNew.Activate

    ' Save to Downloads folder
    savePath = Environ("USERPROFILE") & "\Downloads\"
    fileName = "Accruals_Report_" & Format(cutoffDate, "YYYY-MM-DD") & ".xlsx"

    ' Check if file already exists and prompt to overwrite
    If Dir(savePath & fileName) <> "" Then
        If MsgBox("File already exists. Overwrite?" & vbCrLf & savePath & fileName, vbYesNo + vbQuestion) = vbNo Then
            RestoreThinking
            MsgBox "Report generation cancelled. The workbook remains open but not saved.", vbInformation
            Exit Sub
        End If

        ' Kill raises if the file is open in another program, which for a report in Downloads is the
        ' likely case. Report that plainly instead of failing on the SaveAs a moment later.
        On Error Resume Next
        Kill savePath & fileName
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo Failed
            RestoreThinking
            MsgBox "Could not replace the existing report -- it is probably open:" & vbCrLf & _
                   savePath & fileName & vbCrLf & vbCrLf & _
                   "Close it and run this again. The new report is still open and unsaved.", _
                   vbExclamation, "Accruals Report"
            Exit Sub
        End If
        On Error GoTo Failed
    End If

    ' Save the workbook
    wbNew.SaveAs fileName:=savePath & fileName, FileFormat:=xlOpenXMLWorkbook

    RestoreThinking

    ' Display summary message
    MsgBox "Accruals report generated successfully!" & vbCrLf & vbCrLf & _
           "13-week records: " & matchCount & vbCrLf & _
           "YTD records: " & matchCountYTD & vbCrLf & _
           "Saved to: " & savePath & fileName, vbInformation, "Report Complete"

    Exit Sub

Failed:
    ' Capture the description before cleanup: restoring state can reset Err.
    Dim failureDetail As String
    failureDetail = Err.Description

    ' Force Excel back to a usable state, then report. The generated workbook is deliberately left
    ' open rather than discarded -- if the failure was in SaveAs, the data is still there to save by
    ' hand.
    ResetThinking
    ResetProtection

    MsgBox "GenerateAccrualsReport stopped: " & failureDetail & vbCrLf & vbCrLf & _
           "Calculation and screen updating have been restored. Any report workbook that was " & _
           "created is still open and unsaved.", vbCritical
End Sub

' Asks which closed fiscal month to cut off at and returns that month's last day once the user confirms the
' date. Returns 0 when the user cancels.
Private Function AskForCutoffDate() As Date
    Dim lastClosedDay As Date, fy As Long, lastMonth As Long
    lastClosedDay = AccrualLastClosedDay(Date)
    fy = FiscalYearOf(lastClosedDay)
    lastMonth = FiscalMonthOf(lastClosedDay)

    Dim answer As String, chosen As Long, cutoff As Date
    Do
        answer = AskForCutoffMonth(fy, lastMonth)
        If Len(answer) = 0 Then Exit Function

        chosen = AccrualMonthFromAnswer(answer, lastMonth)
        If chosen = 0 Then
            MsgBox "'" & answer & "' is not one of the listed months.", vbExclamation, "Accruals Report"
        Else
            cutoff = AccrualMonthEnd(fy, chosen)
            If MsgBox("Fiscal " & AccrualFiscalMonthName(chosen) & " " & fy & " ends " & _
                      Format$(cutoff, "dddd, mm/dd/yyyy") & "." & vbCrLf & vbCrLf & _
                      "Use this as the cut-off date?", vbYesNo + vbQuestion, "Accruals Report - Confirm Cut-off") = vbYes Then
                AskForCutoffDate = cutoff
                Exit Function
            End If
        End If
    Loop
End Function

' The user's answer to "which month": what they clicked or typed, trimmed, or "" if they cancelled.
' Offers the variant's clickable picker first and falls back to a plain prompt when the variant has
' none (PickCutoffMonth returns False), the same split InputSheetName uses for sheets.
Private Function AskForCutoffMonth(ByVal fiscalYear As Long, ByVal lastMonth As Long) As String
    Dim names(0 To 11) As Variant, m As Long
    For m = 1 To 12
        names(m - 1) = AccrualFiscalMonthName(m)
    Next m

    Dim answer As String
    If PickCutoffMonth(fiscalYear, lastMonth, names, answer) Then
        AskForCutoffMonth = VBA.Trim$(answer)
        Exit Function
    End If

    Dim menu As String
    For m = 1 To lastMonth
        menu = menu & vbCrLf & "    " & m & ".  " & names(m - 1)
    Next m
    AskForCutoffMonth = VBA.Trim$(InputBox("Choose the cut-off month for fiscal " & fiscalYear & _
                                           " (type its number or name):" & vbCrLf & menu, _
                                           "Accruals Report - Cut-off Month"))
End Function

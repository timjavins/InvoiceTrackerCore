' Fiscal-month arithmetic behind the accruals report's cut-off picker (GenerateAccrualsReport.vb).
'
' Kept apart from the report so it is pure computation over FiscalCalendar.vb: no worksheet, no
' tenant config, no dialog. That is what lets tests/Test-AccrualCutoff.ps1 exercise it without
' stubbing the report's dependencies.
'
' Fiscal month indexes are 1-12 where 1 = February and 12 = January, as in FiscalCalendar.vb.

' Fiscal month names in fiscal order.
Public Function AccrualFiscalMonthName(ByVal fiscalMonth As Long) As String
    If fiscalMonth < 1 Or fiscalMonth > 12 Then Exit Function
    AccrualFiscalMonthName = Choose(fiscalMonth, "February", "March", "April", "May", "June", "July", _
                                    "August", "September", "October", "November", "December", "January")
End Function

' Last day (a Saturday) of a fiscal month.
Public Function AccrualMonthEnd(ByVal fiscalYear As Long, ByVal fiscalMonth As Long) As Date
    AccrualMonthEnd = FiscalMonthStart(fiscalYear, fiscalMonth) + FiscalMonthWeeks(fiscalYear, fiscalMonth) * 7 - 1
End Function

' Last day of the newest fiscal month that has finished before asOf. An accrual cut-off only makes sense
' for a closed month, and in the first weeks of a fiscal year that month is the previous year's January,
' so the caller takes the fiscal year from this date rather than from asOf.
Public Function AccrualLastClosedDay(ByVal asOf As Date) As Date
    AccrualLastClosedDay = FiscalMonthStart(FiscalYearOf(asOf), FiscalMonthOf(asOf)) - 1
End Function

' Reads what the user typed at the month prompt: a list number or a month name. Returns the fiscal month
' (1-12) or 0 when the text is not one of months 1 to lastMonth.
Public Function AccrualMonthFromAnswer(ByVal answer As String, ByVal lastMonth As Long) As Long
    answer = Trim$(answer)
    If Len(answer) = 0 Then Exit Function

    Dim m As Long
    If IsNumeric(answer) Then
        Dim n As Double
        n = CDbl(answer)
        If n = Int(n) And n >= 1 And n <= lastMonth Then AccrualMonthFromAnswer = CLng(n)
        Exit Function
    End If

    For m = 1 To lastMonth
        If StrComp(answer, AccrualFiscalMonthName(m), vbTextCompare) = 0 Then
            AccrualMonthFromAnswer = m
            Exit Function
        End If
    Next m
End Function

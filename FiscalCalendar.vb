' Nordstrom 4-5-4 retail fiscal calendar, computed from one anchor rule.
'
' Pure computation: no worksheet, no workbook, no file, no network. That is deliberate --
' the calendar used to be looked up in a Finance-owned SharePoint workbook, which meant every
' consumer inherited that file's availability and layout. The rule below reproduces all 46
' published years exactly, so nothing needs to read it at run time.
'
' The rule:
'   1. Fiscal year N ends on the Saturday closest to 31 January of year N+1. Year N+1 starts
'      the Sunday after. Weeks start Sunday; years end Saturday.
'   2. Each quarter is 4,5,4 weeks -- 52 weeks in a normal year.
'   3. A year runs 53 weeks exactly when the gap between consecutive year-end anchors is 371
'      days rather than 364. The extra week lands on the final fiscal month, January.
'
' Rules 1 and 3 absorb leap-year drift at the anchor, which is why nothing here reasons about
' February directly.
'
' Verified against tests/../fiscal-calendar-fixture.csv (FY2005-FY2050) by
' tests/Test-FiscalCalendar.ps1. See ADR and the design doc referenced in that plan.
'
' Fiscal month indexes are 1-12 where 1 = February and 12 = January, because that is the order
' the fiscal year runs in. Callers wanting a calendar month name must map it themselves.

' Proves the module loaded and is callable. Used by the test harness smoke check.
Public Function FiscalCalendarSelfCheck() As String
    FiscalCalendarSelfCheck = "ok"
End Function

' The Saturday nearest a given date. Ties cannot occur: a date is at most 3 days from one
' Saturday and at least 4 from the other, so "forward if 3 or fewer, else back" is total.
Public Function SaturdayClosestTo(ByVal anchor As Date) As Date
    Dim dow As Long
    dow = Weekday(anchor, vbSunday)      ' 1 = Sunday ... 7 = Saturday

    Dim forwardDays As Long
    forwardDays = 7 - dow                ' 0 when anchor is already Saturday

    If forwardDays <= 3 Then
        SaturdayClosestTo = anchor + forwardDays
    Else
        SaturdayClosestTo = anchor - dow ' the previous Saturday
    End If
End Function

' Last day of the fiscal year: the Saturday closest to 31 January of the FOLLOWING calendar year.
Public Function FiscalYearEnd(ByVal fiscalYear As Long) As Date
    FiscalYearEnd = SaturdayClosestTo(DateSerial(fiscalYear + 1, 1, 31))
End Function

' First day of the fiscal year -- the Sunday after the previous year ended. Derived from the
' previous year's anchor rather than from this year's, so the two can never disagree about
' where the boundary is.
Public Function FiscalYearStart(ByVal fiscalYear As Long) As Date
    FiscalYearStart = FiscalYearEnd(fiscalYear - 1) + 1
End Function

' 52 or 53. A year is long exactly when its anchor-to-anchor span is 371 days rather than 364 --
' which is how the leap-year drift the anchor absorbs becomes visible as a whole extra week.
'
' Deliberately NOT a periodicity rule. The long years are 2006, 2012, 2017, 2023, 2028, 2034,
' 2040, 2045 -- gaps of 6,5,6,5,6,6,5. That is not a 5-6 alternation (2028 -> 2034 -> 2040 is two
' consecutive sixes), so anything counting years since the last long one is wrong twice this century.
Public Function FiscalYearWeeks(ByVal fiscalYear As Long) As Long
    Dim spanDays As Long
    spanDays = CLng(FiscalYearEnd(fiscalYear) - FiscalYearEnd(fiscalYear - 1))
    ' CLng rounds rather than truncates; this is safe only because FiscalYearEnd returns DateSerial values with no time component.
    FiscalYearWeeks = spanDays \ 7
End Function

' Weeks in a fiscal month. fiscalMonth is 1-12 where 1 = February, 12 = January.
'
' The cycle is 4,5,4 per quarter, so the 5-week months are the second of each quarter --
' fiscal months 2, 5, 8, 11, i.e. those where fiscalMonth Mod 3 = 2. In a 53-week year the
' extra week lands on January (month 12), making it 5 and its quarter 14 weeks.
Public Function FiscalMonthWeeks(ByVal fiscalYear As Long, ByVal fiscalMonth As Long) As Long
    If fiscalMonth < 1 Or fiscalMonth > 12 Then
        Err.Raise 5, "FiscalMonthWeeks", _
            "fiscalMonth must be 1-12 (1 = February, 12 = January); got " & fiscalMonth & "."
    End If

    Dim weeks As Long
    If fiscalMonth Mod 3 = 2 Then
        weeks = 5
    Else
        weeks = 4
    End If

    If fiscalMonth = 12 Then
        If FiscalYearWeeks(fiscalYear) = 53 Then weeks = weeks + 1
    End If

    FiscalMonthWeeks = weeks
End Function

' First day (a Sunday) of a fiscal month.
Public Function FiscalMonthStart(ByVal fiscalYear As Long, ByVal fiscalMonth As Long) As Date
    If fiscalMonth < 1 Or fiscalMonth > 12 Then
        Err.Raise 5, "FiscalMonthStart", _
            "fiscalMonth must be 1-12 (1 = February, 12 = January); got " & fiscalMonth & "."
    End If

    Dim weeksBefore As Long
    Dim i As Long
    For i = 1 To fiscalMonth - 1
        weeksBefore = weeksBefore + FiscalMonthWeeks(fiscalYear, i)
    Next i

    FiscalMonthStart = FiscalYearStart(fiscalYear) + (weeksBefore * 7)
End Function

' First day (a Sunday) of a fiscal week. weekNumber is 1-based and must fall inside the year,
' which is 52 or 53 weeks long -- asking for week 53 of a 52-week year is a caller bug, not a
' value to guess at.
Public Function FiscalWeekStart(ByVal fiscalYear As Long, ByVal weekNumber As Long) As Date
    Dim total As Long
    total = FiscalYearWeeks(fiscalYear)

    If weekNumber < 1 Or weekNumber > total Then
        Err.Raise 5, "FiscalWeekStart", _
            "weekNumber must be 1-" & total & " for FY" & fiscalYear & "; got " & weekNumber & "."
    End If

    FiscalWeekStart = FiscalYearStart(fiscalYear) + ((weekNumber - 1) * 7)
End Function

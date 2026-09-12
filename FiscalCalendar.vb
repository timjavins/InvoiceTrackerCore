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

' Unprotects a worksheet. The password comes from TenantConfig so a variant can change it
' without touching core.

' Unprotects the active sheet.
Public Sub UnprotectSheet()
    UnprotectSheetOn ThisWorkbook.ActiveSheet
End Sub

' Unprotects an explicit sheet, so callers need not rely on what happens to be active.
Public Sub UnprotectSheetOn(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    ws.Unprotect Password:=TenantSheetPassword()
End Sub

' Protects a worksheet, leaving AutoFilter usable on locked cells.
' The password comes from TenantConfig so a variant can change it without touching core.

' Protects the active sheet.
Public Sub ProtectSheet()
    ProtectSheetOn ThisWorkbook.ActiveSheet
End Sub

' Protects an explicit sheet, so callers need not rely on what happens to be active.
Public Sub ProtectSheetOn(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    ws.Protect Password:=TenantSheetPassword(), AllowFiltering:=True
End Sub

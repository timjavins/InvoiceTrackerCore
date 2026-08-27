' Unprotects a worksheet. The password comes from TenantConfig so a variant can change it
' without touching core.
'
' Opens a nesting-counted unprotect window -- see ProtectSheet.vb, which holds the depth and
' found-state maps and explains why the counting exists. Always pair with ProtectSheetOn.

' Unprotects the active sheet.
Public Sub UnprotectSheet()
    UnprotectSheetOn ThisWorkbook.ActiveSheet
End Sub

' Opens an unprotect window on an explicit sheet, so callers need not rely on what happens to
' be active. Nested calls deepen the window rather than reopening it.
Public Sub UnprotectSheetOn(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    EnsureProtectMaps

    Dim k As String
    k = ProtectionKey(ws)

    Dim depth As Long
    If protectDepth.Exists(k) Then depth = CLng(protectDepth(k))

    ' Only the outermost window records what to go back to. An inner call must not, because
    ' by then the sheet is already unprotected and capturing that would make the outermost
    ' ProtectSheetOn "restore" the sheet to open.
    If depth = 0 Then protectPrior(k) = ws.ProtectContents

    protectDepth(k) = depth + 1

    If ws.ProtectContents Then ws.Unprotect Password:=TenantSheetPassword()
End Sub

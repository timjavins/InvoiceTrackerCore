' Protects a worksheet, leaving AutoFilter usable on locked cells.
' The password comes from TenantConfig so a variant can change it without touching core.
'
' Protection is nesting-counted, per sheet, for the same reason PauseThinking is: a helper
' that balances its own Unprotect/Protect pair used to re-protect the sheet out from under a
' caller that was still mid-operation. WriteFormulas_Tracker did exactly that, and the next
' post-import step -- PopulateInvoiceTypeForNewRows, which has no handler of its own -- died
' on a protected-sheet write and took the whole import with it.
'
' Per sheet rather than workbook-wide because UpdateSearchValues holds a window open on the
' Helper sheet while the tracker's own state is independent.
'
' The depth map and the found-state map live here; UnprotectSheet.vb shares them, since the
' assembler concatenates every module into one.

' "workbook|sheet" -> how many nested unprotect windows are currently open on it.
Private protectDepth As Object

' "workbook|sheet" -> whether the sheet was protected when the outermost window opened, so
' closing that window puts it back the way it was found rather than guessing.
Private protectPrior As Object

' Protects the active sheet.
Public Sub ProtectSheet()
    ProtectSheetOn ThisWorkbook.ActiveSheet
End Sub

' Closes one unprotect window on an explicit sheet, so callers need not rely on what happens
' to be active. Only the outermost window actually re-protects.
Public Sub ProtectSheetOn(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    EnsureProtectMaps

    Dim k As String
    k = ProtectionKey(ws)

    Dim depth As Long
    If protectDepth.Exists(k) Then depth = CLng(protectDepth(k))

    ' No matching unprotect. Protect anyway: every caller written before the counter existed
    ' relied on ProtectSheetOn being unconditional, and refusing here would silently leave a
    ' tracker unprotected.
    If depth = 0 Then
        ApplyProtection ws
        Exit Sub
    End If

    depth = depth - 1
    protectDepth(k) = depth

    ' An inner helper finishing does not end the caller's window.
    If depth > 0 Then Exit Sub

    Dim wasProtected As Boolean
    wasProtected = True
    If protectPrior.Exists(k) Then wasProtected = CBool(protectPrior(k))

    ' Sheets that were already unprotected when we arrived stay that way. 'BU List',
    ' 'Store Directory' and 'Menu' are deliberately open, and newly protecting them would be
    ' a behaviour change no caller asked for.
    If wasProtected Then ApplyProtection ws

    protectDepth.Remove k
    If protectPrior.Exists(k) Then protectPrior.Remove k
End Sub

' Force every open unprotect window closed and put each sheet back the way it was found.
'
' For error handlers, and for orchestration entry points to call before opening the first
' window -- the same role ResetThinking plays for PauseThinking. Without it, an abort inside
' a nested window strands the depth above zero, and every later UnprotectSheetOn on that
' sheet becomes a silent no-op that fails the next run the same way with no clue why.
Public Sub ResetProtection()
    If protectDepth Is Nothing Then Exit Sub
    If protectPrior Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Dim k As String
    For Each ws In ThisWorkbook.Worksheets
        k = ProtectionKey(ws)
        If protectPrior.Exists(k) Then
            If CBool(protectPrior(k)) Then ApplyProtection ws
        End If
    Next ws

    protectDepth.RemoveAll
    protectPrior.RemoveAll
End Sub

' Both maps are created lazily: the stacked code lands in ThisWorkbook, a class module,
' where a module-level initialiser is not available.
Private Sub EnsureProtectMaps()
    If protectDepth Is Nothing Then Set protectDepth = CreateObject("Scripting.Dictionary")
    If protectPrior Is Nothing Then Set protectPrior = CreateObject("Scripting.Dictionary")
End Sub

' Sheet identity has to include the workbook: both tracker workbooks are routinely open in
' the same Excel instance and their sheet names collide.
Private Function ProtectionKey(ByVal ws As Worksheet) As String
    ProtectionKey = ws.Parent.Name & "|" & ws.Name
End Function

Private Sub ApplyProtection(ByVal ws As Worksheet)
    If ws.ProtectContents Then Exit Sub
    ws.Protect Password:=TenantSheetPassword(), AllowFiltering:=True
End Sub

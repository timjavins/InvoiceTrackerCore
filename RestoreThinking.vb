' Restores what PauseThinking suspended. The saved-state variables live in
' PauseThinking.vb; the assembler concatenates both into a single module, so they are
' in scope here.

Public Sub RestoreThinking()
    ' An unbalanced call -- nothing was paused, or a caller suspended settings by hand without going
    ' through PauseThinking -- must not assign the saved-state variables below, because they were
    ' never captured. savedCalculation would be 0, which is not a valid XlCalculation and raises;
    ' inside an On Error Resume Next cleanup handler that raise is swallowed and Excel is left
    ' suspended with nothing to show why. Fall back to the known-good all-on state instead.
    If pauseDepth = 0 Then
        ResetThinking
        Exit Sub
    End If

    pauseDepth = pauseDepth - 1

    ' Unwind only when the outermost pause completes.
    If pauseDepth > 0 Then Exit Sub

    Application.Calculation = savedCalculation
    Application.ScreenUpdating = savedScreenUpdating
    Application.EnableEvents = savedEnableEvents
    Application.DisplayAlerts = savedDisplayAlerts
    Application.AutoRecover.Enabled = savedAutoRecover
End Sub

' Whether a pause is currently active.
Public Function ThinkingIsPaused() As Boolean
    ThinkingIsPaused = (pauseDepth > 0)
End Function

' Lifts a pause so a modal dialog can paint, reporting whether it actually lifted one. Pair with
' ResumeThinkingAfterDialog, passing the returned value.
'
' A modal UserForm shown while ScreenUpdating is off paints as a blank grey rectangle. Callers
' cannot simply call RestoreThinking themselves: when nothing is paused, the saved-state variables
' below have never been captured, so restoring would assign Calculation = 0, which is not a valid
' XlCalculation and raises. This makes the unwind safe to call unconditionally.
Public Function SuspendThinkingForDialog() As Boolean
    If pauseDepth = 0 Then Exit Function
    RestoreThinking
    SuspendThinkingForDialog = True
End Function

' Restores the pause that SuspendThinkingForDialog lifted, if it lifted one.
Public Sub ResumeThinkingAfterDialog(ByVal wasPaused As Boolean)
    If wasPaused Then PauseThinking
End Sub

' Force a full restore regardless of nesting. For error handlers that need to be certain
' Excel is usable again after an unbalanced Pause.
Public Sub ResetThinking()
    pauseDepth = 0
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.AutoRecover.Enabled = True
End Sub

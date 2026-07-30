' Restores what PauseThinking suspended. The saved-state variables live in
' PauseThinking.vb; the assembler concatenates both into a single module, so they are
' in scope here.

Public Sub RestoreThinking()
    If pauseDepth > 0 Then
        pauseDepth = pauseDepth - 1
    End If

    ' Unwind only when the outermost pause completes.
    If pauseDepth > 0 Then Exit Sub

    Application.Calculation = savedCalculation
    Application.ScreenUpdating = savedScreenUpdating
    Application.EnableEvents = savedEnableEvents
    Application.DisplayAlerts = savedDisplayAlerts
    Application.AutoRecover.Enabled = savedAutoRecover
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

' Suspends the Excel features that make bulk sheet operations slow. Always pair with
' RestoreThinking (see RestoreThinking.vb, which shares the saved-state variables below
' because the assembler concatenates every module into one).
'
' Grafted from both variants: Securitas suspended five settings but restored them to
' hardcoded defaults (clobbering whatever the user had); JCI restored the caller's
' actual prior state but only covered two settings. This keeps the wider coverage and
' the correct restore semantics.
'
' Nesting is counted, so an outer PauseThinking is not undone early by an inner
' routine's RestoreThinking.

Private pauseDepth As Long
Private savedCalculation As XlCalculation
Private savedScreenUpdating As Boolean
Private savedEnableEvents As Boolean
Private savedDisplayAlerts As Boolean
Private savedAutoRecover As Boolean

Public Sub PauseThinking()
    ' Only the outermost call captures state; inner calls just deepen the nesting.
    If pauseDepth = 0 Then
        savedCalculation = Application.Calculation
        savedScreenUpdating = Application.ScreenUpdating
        savedEnableEvents = Application.EnableEvents
        savedDisplayAlerts = Application.DisplayAlerts
        savedAutoRecover = Application.AutoRecover.Enabled
    End If

    pauseDepth = pauseDepth + 1

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.AutoRecover.Enabled = False   ' AutoRecover, not AutoSave
End Sub

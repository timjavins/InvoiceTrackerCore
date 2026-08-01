' Suspends the Excel features that make bulk sheet operations slow. Always pair with
' RestoreThinking (see RestoreThinking.vb, which shares the saved-state variables below
' because the assembler concatenates every module into one).
'
' Suspends five settings -- Securitas's coverage -- and restores them to the on state.
'
' An earlier version saved and restored the caller's actual prior values, which read better
' but was wrong in practice: if a run ended without restoring, the next PauseThinking captured
' the already-suspended values and then faithfully "restored" calculation to Manual. No error,
' no clue, and self-perpetuating. Excel's normal state is all five on, and a workbook stuck
' with calculation off is far worse than losing a deliberately non-default setting, so the
' restore targets are fixed rather than captured.
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
        ' Capture what to go back to -- but never capture a suspended state. If a previous
        ' run ended without restoring, Application.Calculation is already xlCalculationManual
        ' and screen updating is already off; saving those would "restore" Excel to broken,
        ' with no error to show why. That is self-perpetuating: once off, always off.
        If Application.Calculation = xlCalculationManual Then
            savedCalculation = xlCalculationAutomatic
        Else
            savedCalculation = Application.Calculation
        End If

        savedScreenUpdating = True
        savedEnableEvents = True
        savedDisplayAlerts = True
        savedAutoRecover = True
    End If

    pauseDepth = pauseDepth + 1

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.AutoRecover.Enabled = False   ' AutoRecover, not AutoSave
End Sub

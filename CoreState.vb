' Module-level state shared by core routines.
'
' Declared here rather than in a variant's Header.vb so that any variant consuming core
' gets it automatically. VBA allows module-level declarations anywhere in a module, so
' this file's position in the stack does not matter -- unlike Option Explicit, which must
' precede all procedures and therefore belongs only in Header.vb.

' Set by a core routine when it could not complete an operation, so the caller can decide
' whether to continue. Callers are responsible for clearing it before a run.
Public failureState As Boolean

' Set when a run completed but with recoverable problems worth reporting at the end.
Public withErrors As Boolean

' Clears the run-level error flags. Call at the start of an orchestration entry point.
Public Sub ResetCoreState()
    failureState = False
    withErrors = False
End Sub

' Module-level state shared by core routines.
'
' Declared here rather than in a variant's Header.vb so that any variant consuming core
' gets it automatically.
'
' VBA requires EVERY module-level declaration to precede the first procedure -- not just
' Option Explicit. Anything after one raises "Only comments may appear after End Sub, End
' Function, or End Property". A module in the middle of the stack cannot satisfy that on its
' own, so Stack-VBFiles.ps1 hoists declarations into a header block at the top of the
' assembled file. Declaring state in a core module is therefore fine; the assembler places it
' correctly.
'
' The stacked code is pasted into ThisWorkbook, which is a class module, so avoid the
' declarations VBA forbids as public members of an object module: Public Const, public fixed
' size arrays, fixed-length strings, and Declare. Private Const is fine.

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

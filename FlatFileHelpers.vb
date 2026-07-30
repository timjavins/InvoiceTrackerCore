' Value predicates used when building a Coupa upload flat file. Kept separate from the
' generators themselves, which remain per-variant -- see the note at the bottom.

' Whether a business-unit value is the tenant's credit BU (Securitas: 302). A tenant with
' no separate credit identity declares an empty TenantCreditBusinessUnit and never matches.
Public Function IsCreditBusinessUnit(ByVal buValue As Variant) As Boolean
    Dim creditBU As String
    creditBU = Trim$(TenantCreditBusinessUnit())
    If Len(creditBU) = 0 Then Exit Function

    If IsError(buValue) Or IsNull(buValue) Then Exit Function

    Dim buText As String
    buText = VBA.Trim$(CStr(buValue))
    If Len(buText) = 0 Then Exit Function

    ' Compare numerically when both sides are numbers, so "302" and 302.0 both match.
    If VBA.IsNumeric(buText) And VBA.IsNumeric(creditBU) Then
        IsCreditBusinessUnit = (CLng(CDbl(buText)) = CLng(CDbl(creditBU)))
    Else
        IsCreditBusinessUnit = (StrComp(buText, creditBU, vbTextCompare) = 0)
    End If
End Function

' Whether a requisition status means "not yet raised", so the row still needs a flat file.
' Blank, missing, and zero all count as pending.
Public Function IsReqStatusPending(ByVal value As Variant) As Boolean
    If IsError(value) Or IsNull(value) Then
        IsReqStatusPending = True
        Exit Function
    End If

    Dim statusText As String
    statusText = VBA.Trim$(CStr(value))

    If Len(statusText) = 0 Then
        IsReqStatusPending = True
    ElseIf VBA.IsNumeric(statusText) Then
        IsReqStatusPending = (CDbl(statusText) = 0)
    Else
        IsReqStatusPending = False
    End If
End Function

' Whether a cell value can be treated as a number, tolerating errors and nulls.
'
' Named distinctly from VBA's own IsNumeric on purpose: Securitas defined a Private
' IsNumeric that shadowed the built-in within its module. In a concatenated stack that
' shadowing would apply to every module, so the helper is renamed rather than moved as-is.
Public Function IsNumericValue(ByVal value As Variant) As Boolean
    If IsError(value) Then Exit Function
    If IsNull(value) Then Exit Function

    If VBA.VarType(value) = vbString Then
        On Error Resume Next
        IsNumericValue = Not VBA.IsError(CDbl(value))
        On Error GoTo 0
    Else
        IsNumericValue = VBA.IsNumeric(value)
    End If
End Function

' NOTE ON SCOPE
'
' The flat file GENERATORS stay per-variant. They are not one behaviour with different
' constants: Securitas's is ~730 lines across 8 procedures with bill-code grouping, credit
' (BU 302) handling, multi-store cost allocation and monitoring rules; JCI's is ~230 lines
' in a single procedure with none of those. Merging them would be a rewrite of JCI's output
' format, not an extraction, and the output feeds a real Coupa upload.

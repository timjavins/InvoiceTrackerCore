' Value predicates used when building a Coupa upload flat file. Kept separate from the
' generators themselves, which remain per-variant -- see the note at the bottom.

' Whether a business-unit value is the tenant's credit BU (TenantA: 302). A tenant with
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

' The Coupa chart-of-accounts name a requisition line with this business unit posts to.
'
' The instance keeps a separate chart of accounts for the credit entity, and a line must post to the
' one its business unit belongs to: Coupa rejects a line whose account string is not valid in the chart
' it names. Which BU is the credit BU, and what the two charts are called, are facts about the Coupa
' instance, so they are configuration, not code: TenantCreditBusinessUnit and
' TenantChartOfAccountName(isCredit). Every tenant declares both. A tenant that leaves the BU empty
' never matches and so always posts to its normal chart -- right only if it genuinely never bills the
' credit entity, so declare them even when no credit row has been seen yet. See ADR-0006.
Public Function ChartOfAccountNameForBU(ByVal buValue As Variant) As String
    ChartOfAccountNameForBU = TenantChartOfAccountName(IsCreditBusinessUnit(buValue))
End Function

' Whether a supplier's Coupa site name agrees with the chart of accounts a line posts to. A site name
' ends in the name of the chart it belongs to ("<NAME>_<ADDRESS>_<CITY>_<chart name>"), and a site for
' one chart on a line that posts to the other is rejected on upload. Some names wrap the whole site in
' parentheses, "<NAME> (<NAME>_<ADDRESS>_<CITY>_<chart name>)", so one trailing ")" is ignored.
'
' False only on that contradiction. A name that ends in neither chart's name -- a placeholder such as
' "ODP -", or a blank -- gives no verdict and returns True, as does a tenant that configures a single
' chart: this is a guard against the one known mismatch, not a validator of supplier names, and a blank
' supplier is reported by its own caller.
Public Function SupplierSiteMatchesChart(ByVal supplierName As String, ByVal chartName As String) As Boolean
    SupplierSiteMatchesChart = True

    Dim nm As String
    nm = LCase$(Trim$(supplierName))
    If Right$(nm, 1) = ")" Then nm = Trim$(Left$(nm, Len(nm) - 1))

    Dim chart As String, normalChart As String, creditChart As String, other As String
    chart = LCase$(Trim$(chartName))
    normalChart = LCase$(Trim$(TenantChartOfAccountName(False)))
    creditChart = LCase$(Trim$(TenantChartOfAccountName(True)))

    If Len(chart) = 0 Then Exit Function
    If chart = normalChart Then
        other = creditChart
    ElseIf chart = creditChart Then
        other = normalChart
    End If
    If Len(other) = 0 Or other = chart Then Exit Function

    If FlatFileSiteEndsWith(nm, chart) Then Exit Function
    If FlatFileSiteEndsWith(nm, other) Then SupplierSiteMatchesChart = False
End Function

Private Function FlatFileSiteEndsWith(ByVal lowerName As String, ByVal lowerChart As String) As Boolean
    FlatFileSiteEndsWith = (Right$(lowerName, Len(lowerChart) + 1) = "_" & lowerChart)
End Function

' Why a requisition line cannot be uploaded as built, in words a person can act on, or "" when it can.
' Every flat-file generator calls this once the row's business unit and its supplier are known, and
' skips the row with the reason when it is not empty -- Coupa rejects these lines, and a rejected line
' is invisible until someone reconciles against Coupa days later.
'
' Two things are checked. (1) A credit-BU line needs a supplier with a credit identity: not every
' supplier has one. A tenant, or a sheet, that has none returns 0 or "" or "N/A" for the number and
' "" for the name, and this says so instead of letting the line fall back to the normal supplier.
' (2) The supplier's site must belong to the chart the line posts to (SupplierSiteMatchesChart).
'
' A normal (non-credit) line with no supplier yet is not this function's concern: the caller already
' reports a supplier it could not resolve.
Public Function LineIdentityProblem(ByVal buValue As Variant, ByVal supplierName As String, _
                                    ByVal supplierNumber As Variant) As String
    Dim chart As String
    chart = ChartOfAccountNameForBU(buValue)

    If IsCreditBusinessUnit(buValue) Then
        If Not FlatFileSupplierNumberIsReal(supplierNumber) Or Len(Trim$(supplierName)) = 0 Then
            LineIdentityProblem = "business unit " & Trim$(CStr(buValue)) & " posts to the " & chart & _
                " chart of accounts, but this supplier has no " & chart & " supplier number. Set one up in Coupa, " & _
                "or bill this store another way."
            Exit Function
        End If
    End If

    If Not SupplierSiteMatchesChart(supplierName, chart) Then
        LineIdentityProblem = "business unit " & Trim$(CStr(buValue)) & " posts to the " & chart & _
            " chart of accounts, but the supplier site (" & Trim$(supplierName) & _
            ") belongs to the other chart of accounts."
    End If
End Function

' A supplier number that names a real Coupa supplier: not blank, not "N/A", not zero.
Private Function FlatFileSupplierNumberIsReal(ByVal v As Variant) As Boolean
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function

    Dim t As String
    t = Trim$(CStr(v))
    If Len(t) = 0 Then Exit Function
    If StrComp(t, "N/A", vbTextCompare) = 0 Then Exit Function
    If IsNumeric(t) Then
        If CDbl(t) = 0 Then Exit Function
    End If
    FlatFileSupplierNumberIsReal = True
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
' Named distinctly from VBA's own IsNumeric on purpose: TenantA defined a Private
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
' constants: TenantA's is ~730 lines across 8 procedures with bill-code grouping, credit
' (BU 302) handling, multi-store cost allocation and monitoring rules; TenantB's is ~230 lines
' in a single procedure with none of those. Merging them would be a rewrite of TenantB's output
' format, not an extraction, and the output feeds a real Coupa upload.
'
' What does come out from under them is whole decisions both generators have to make the same
' way. Requester resolution was the fourth, after the three predicates above -- it now lives in
' RequesterEmail.vb and CoupaUsers.vb. See ADR-0005. The fifth is which chart of accounts a row
' posts to, above. That is the pattern to follow: extract a decision, not the generator around it.

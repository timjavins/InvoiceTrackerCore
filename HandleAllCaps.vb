' Converts an all-caps name to proper case, leaving mixed-case names alone.
' Coupa returns some approver names in caps, e.g. "JANE SMITH" -> "Jane Smith".
Public Function HandleAllCaps(ByVal name As String) As String
    If name = VBA.UCase$(name) Then
        HandleAllCaps = StrConv(VBA.LCase$(name), vbProperCase)
    Else
        HandleAllCaps = name
    End If
End Function

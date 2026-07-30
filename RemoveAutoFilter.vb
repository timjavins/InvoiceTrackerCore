' Removes a worksheet's AutoFilter and reports which row it was on, so a caller can put
' it back afterwards with RestoreAutoFilter.
'
' Returns 0 when no AutoFilter was applied.
Public Function RemoveAutoFilter(ByVal wsTarget As Worksheet) As Long
    If wsTarget Is Nothing Then Exit Function

    Dim filterRow As Long

    On Error Resume Next

    If wsTarget.AutoFilterMode Then
        filterRow = wsTarget.AutoFilter.Range.Row
        wsTarget.AutoFilterMode = False
        RemoveAutoFilter = filterRow
    Else
        RemoveAutoFilter = 0
    End If

    On Error GoTo 0
End Function

Function MapSourceToTargetColumns(sourceColIndexes As Object, targetHeaderDict As Object) As Object
    Dim colTarget As Long
    Dim currentCol As Variant
    Dim targetMap As Object
    Set targetMap = CreateObject("Scripting.Dictionary")
    ' Loop through sourceColIndexes and find the matching columns in the target worksheet
    For Each currentCol In sourceColIndexes.Keys
        If targetHeaderDict.Exists(currentCol) Then
            ' Map the source column to the target column
            targetMap.Add currentCol, targetHeaderDict(currentCol)
        Else
            colTarget = InputTargetColumnName(VBA.CStr(currentCol), targetHeaderDict)
            If failureState Then
                withErrors = True
                failureState = False ' Reset the failure state
                MsgBox "The target column for '" & currentCol & "' is unknown. Exiting process.", vbExclamation
                Exit Function
            End If
            targetMap.Add currentCol, colTarget
        End If
    Next currentCol
    Set MapSourceToTargetColumns = targetMap
End Function
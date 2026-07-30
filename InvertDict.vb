' Swaps a dictionary's keys and values. Used to turn a {column index -> header name} map
' into {header name -> column index}.
'
' Values that repeat would collide, so the first occurrence wins and later ones are
' skipped. The original raised an error in that case, which was easy to trip on a source
' file with duplicate headers.
Public Function InvertDict(ByVal dict As Object) As Object
    Dim inverted As Object
    Set inverted = CreateObject("Scripting.Dictionary")
    Set InvertDict = inverted

    If dict Is Nothing Then Exit Function

    Dim key As Variant
    For Each key In dict.Keys
        If Not inverted.Exists(dict(key)) Then
            inverted.Add dict(key), key
        Else
            Debug.Print "InvertDict: duplicate value '" & CStr(dict(key)) & "' -- keeping first."
        End If
    Next key
End Function

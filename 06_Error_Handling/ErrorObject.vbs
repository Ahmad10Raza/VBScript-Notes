On Error Resume Next
Dim result
result = 10 / 0

If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Err.Clear ' Clear the error
End If
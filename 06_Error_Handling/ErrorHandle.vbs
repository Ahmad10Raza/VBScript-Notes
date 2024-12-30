On Error Resume Next
' Cause multiple errors
Dim file, result
file = "nonexistent_file.txt"
Open file For Input As #1 ' This will cause a file not found error
If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Err.Clear
End If
result = 10 / 0 ' This will cause a division by zero error
If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Err.Clear
End If



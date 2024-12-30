On Error Resume Next
' Code that may cause an error
Dim x, y
x = 10
y = 0
MsgBox x / y

' Check for errors
If Err.Number <> 0 Then
    MsgBox "Error occurred: " & Err.Description
End If

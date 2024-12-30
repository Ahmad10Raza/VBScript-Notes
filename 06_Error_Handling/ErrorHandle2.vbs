On Error Resume Next
' Simulate an error in a function
Call ErrorFunction

If Err.Number <> 0 Then
    MsgBox "Error occurred in: " & Err.Source & vbCrLf & "Error: " & Err.Description
    Err.Clear
End If

Sub ErrorFunction
    Dim x, y
    x = 10
    y = 0
    MsgBox x / y ' Causes division by zero error
End Sub

On Error GoTo ErrorHandler

' Code that may cause an error
Dim x, y
x = 10
y = 0
MsgBox x / y

Exit Sub ' Skip error handler if no error occurs

ErrorHandler:
MsgBox "Error occurred: " & Err.Description

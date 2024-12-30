' Define the Function to Return a String
Function GetGreeting(name)
    GetGreeting = "Hello, " & name & "!"
End Function

' Calling the Function
Dim greeting
greeting = GetGreeting("Alice")
MsgBox greeting

' Define the Function with Optional Parameter
Function GreetUser(name, Optional greeting)
    If IsMissing(greeting) Then
        greeting = "Hello" ' Default greeting
    End If
    GreetUser = greeting & ", " & name & "!"
End Function

' Calling the Function with and without the Optional Parameter
Dim message
message = GreetUser("Alice")
MsgBox message ' Uses default greeting "Hello"

message = GreetUser("Bob", "Good morning")
MsgBox message ' Uses custom greeting "Good morning"    
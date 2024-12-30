' Define the Subroutine with an Optional Parameter
Sub GreetUser(name, Optional greeting)
    ' Check if greeting was provided
    If IsMissing(greeting) Then
        greeting = "Hello" ' Default value if greeting is not provided
    End If
    MsgBox greeting & ", " & name & "!"
End Sub

' Calling the Subroutine with and without the Optional Parameter
GreetUser "Alice" ' Uses default greeting "Hello"
GreetUser "Bob", "Good morning" ' Uses custom greeting "Good morning"

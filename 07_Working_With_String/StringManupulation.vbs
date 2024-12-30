' VBScript String Manipulation Functions

' 1. Len Function
Dim text
text = "Hello, World!"
MsgBox "Length of text: " & Len(text) ' Output: 13

' 2. Mid Function
text = "Hello, World!"
MsgBox "Mid Function: " & Mid(text, 8, 5) ' Output: World

' 3. Left Function
text = "Hello, World!"
MsgBox "Left Function: " & Left(text, 5) ' Output: Hello

' 4. Right Function
text = "Hello, World!"
MsgBox "Right Function: " & Right(text, 6) ' Output: World!

' 5. InStr Function
text = "Hello, World!"
Dim pos
pos = InStr(text, "World")
MsgBox "InStr Function: Position of 'World': " & pos ' Output: 8

' 6. Replace Function
text = "Hello, World!"
text = Replace(text, "World", "VBScript")
MsgBox "Replace Function: " & text ' Output: Hello, VBScript!

' 7. UCase and LCase Functions
text = "Hello, World!"
MsgBox "UCase Function: " & UCase(text) ' Output: HELLO, WORLD!
MsgBox "LCase Function: " & LCase(text) ' Output: hello, world!

' 8. Trim, LTrim, and RTrim Functions
text = "  Hello, World!  "
MsgBox "Trim Function: " & Trim(text) ' Output: Hello, World!
MsgBox "LTrim Function: " & LTrim(text) ' Output: Hello, World!  
MsgBox "RTrim Function: " & RTrim(text) ' Output:   Hello, World!

' 9. Chr and Asc Functions
Dim char, ascii
char = Chr(65) ' Character for ASCII code 65
ascii = Asc("A") ' ASCII value of "A"
MsgBox "Chr Function: " & char ' Output: A
MsgBox "Asc Function: " & ascii ' Output: 65

' 10. StrComp Function
Dim result
result = StrComp("apple", "banana")
If result = 0 Then
    MsgBox "StrComp Function: Strings are equal."
ElseIf result < 0 Then
    MsgBox "StrComp Function: apple is less than banana."
Else
    MsgBox "StrComp Function: apple is greater than banana."
End If

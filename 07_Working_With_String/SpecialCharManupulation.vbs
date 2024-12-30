' VBScript Special Characters Example

' Block 1: Quotation Marks in Strings
Dim text
text = "She said, ""Hello!"""
MsgBox "Block 1: Quotation Marks - " & text ' Output: She said, "Hello!"

' Block 2: Newline (vbCrLf)
text = "Hello" & vbCrLf & "World"
MsgBox "Block 2: Newline - " & text ' Output: Hello (New line) World

' Block 3: Tab (vbTab)
text = "Name:" & vbTab & "John"
MsgBox "Block 3: Tab - " & text ' Output: Name:    John (Tab space between Name: and John)

' Block 4: Carriage Return (vbCr)
text = "Hello" & vbCr & "World"
MsgBox "Block 4: Carriage Return - " & text ' Output: World (Hello is overwritten by World)

' Block 5: Line Feed (vbLf)
text = "Hello" & vbLf & "World"
MsgBox "Block 5: Line Feed - " & text ' Output: Hello (New line) World

' Block 6: Carriage Return and Line Feed (vbCrLf)
text = "Line 1" & vbCrLf & "Line 2"
MsgBox "Block 6: Carriage Return and Line Feed - " & text ' Output: Line 1 (New line) Line 2

' Block 7: File Path (Escaping Backslashes)
Dim path
path = "C:\\Users\\John\\Documents\\File.txt"
MsgBox "Block 7: File Path - " & path ' Output: C:\Users\John\Documents\File.txt

' Block 8: Null String (vbNullString)
Dim empty
empty = vbNullString
MsgBox "Block 8: Null String - " & empty ' Output: (No text)

' Block 9: Special Character Example
text = "Here is a string with special characters:" & vbCrLf & _
       "1. Quotation: ""Hello!""" & vbCrLf & _
       "2. Newline Example:" & vbCrLf & _
       "Line 1" & vbCrLf & "Line 2" & vbCrLf & _
       "3. Tab Example:" & vbTab & "Tabbed Text" & vbCrLf & _
       "4. File path: C:\\Users\\John\\Documents\\File.txt"

MsgBox "Block 9: Special Character Example - " & vbCrLf & text

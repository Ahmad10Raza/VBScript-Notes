Dim fso, file

' Create a FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' Create a text file and write to it
Set file = fso.CreateTextFile("C:\example.txt", True)
file.WriteLine("This is a test file created by VBScript.")
file.Close

' Clean up
Set file = Nothing
Set fso = Nothing

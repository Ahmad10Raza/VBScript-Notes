Dim fso, file
Set fso = CreateObject("Scripting.FileSystemObject")

' Open the file for writing (will overwrite if exists)
Set file = fso.CreateTextFile("C:\example.txt", True)
file.WriteLine("This is a new line of text.")
file.Close

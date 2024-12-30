Dim fso, file
Set fso = CreateObject("Scripting.FileSystemObject")

' Create a new text file
Set file = fso.CreateTextFile("D:\MSLTech\VBScript-Notes\09_Working_With_Files_Folder\example.txt", True)  ' True means to overwrite if exists
file.WriteLine("Hello, this is a test file.")
file.Close

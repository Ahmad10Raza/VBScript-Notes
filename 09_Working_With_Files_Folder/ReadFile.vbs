Dim fso, file, content
Set fso = CreateObject("Scripting.FileSystemObject")

' Open an existing text file for reading
Set file = fso.OpenTextFile("D:\MSLTech\VBScript-Notes\09_Working_With_Files_Folder\example.txt", 1)  ' 1 means reading mode

' Read and display each line of the file
Do While Not file.AtEndOfStream
    content = file.ReadLine
    MsgBox content
Loop
file.Close

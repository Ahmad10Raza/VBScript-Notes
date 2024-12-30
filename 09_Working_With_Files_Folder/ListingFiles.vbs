Dim fso, folder, file
Set fso = CreateObject("Scripting.FileSystemObject")

Set folder = fso.GetFolder("D:\MSLTech\VBScript-Notes\05_Procedures")

' Loop through all files in the folder
For Each file In folder.Files
    MsgBox "File: " & file.Name
Next

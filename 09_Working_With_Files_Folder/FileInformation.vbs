Dim fso, file
Set fso = CreateObject("Scripting.FileSystemObject")

Set file = fso.GetFile("D:\MSLTech\VBScript-Notes\README.md")

' Display file properties
MsgBox "File Name: " & file.Name
MsgBox "File Path: " & file.Path
MsgBox "File Size: " & file.Size & " bytes"
MsgBox "File Date Created: " & file.DateCreated

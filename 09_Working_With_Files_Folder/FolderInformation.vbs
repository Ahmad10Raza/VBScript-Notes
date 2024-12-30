Dim fso, folder
Set fso = CreateObject("Scripting.FileSystemObject")

Set folder = fso.GetFolder("C:\MyFolder")

' Display folder properties
MsgBox "Folder Name: " & folder.Name
MsgBox "Folder Path: " & folder.Path
MsgBox "Folder Size: " & folder.Size & " bytes"
MsgBox "Folder Date Created: " & folder.DateCreated

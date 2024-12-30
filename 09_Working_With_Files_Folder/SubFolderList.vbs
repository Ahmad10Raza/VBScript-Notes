Dim fso, folder, subfolder
Set fso = CreateObject("Scripting.FileSystemObject")

Set folder = fso.GetFolder("C:\MyFolder")

' Loop through all subfolders in the folder
For Each subfolder In folder.Subfolders
    MsgBox "Subfolder: " & subfolder.Name
Next

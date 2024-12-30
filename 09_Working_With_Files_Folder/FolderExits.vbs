Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FolderExists("C:\MyFolder") Then
    MsgBox "Folder exists"
Else
    MsgBox "Folder does not exist"
End If

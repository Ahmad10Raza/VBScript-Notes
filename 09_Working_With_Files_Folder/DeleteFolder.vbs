Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' Delete folder if it exists
If fso.FolderExists("C:\MyFolder") Then
    fso.DeleteFolder("C:\MyFolder")
End If



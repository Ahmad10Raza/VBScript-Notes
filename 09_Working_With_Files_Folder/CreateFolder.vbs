Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' Create a folder
If Not fso.FolderExists("C:\MyFolder") Then
    fso.CreateFolder("C:\MyFolder")
End If

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' Delete the file if it exists
If fso.FileExists("D:\MSLTech\VBScript-Notes\09_Working_With_Files_Folder\example.txt") Then
    fso.DeleteFile("D:\MSLTech\VBScript-Notes\09_Working_With_Files_Folder\example.txt")
End If

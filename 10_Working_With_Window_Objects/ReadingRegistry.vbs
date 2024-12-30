Dim objShell, strKey, strValue

' Create a WScript.Shell object
Set objShell = CreateObject("WScript.Shell")

' Define the registry key and value to read
strKey = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Personal"

' Read the registry value
strValue = objShell.RegRead(strKey)

' Display the value
WScript.Echo "Personal folder path: " & strValue

' Clean up
Set objShell = Nothing

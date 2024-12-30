Dim objShell, strEnvVariable

' Create a WScript.Shell object
Set objShell = CreateObject("WScript.Shell")

' Read the "PATH" environment variable
strEnvVariable = objShell.Environment("System")("PATH")

' Display the environment variable
WScript.Echo "System PATH: " & strEnvVariable

' Clean up
Set objShell = Nothing

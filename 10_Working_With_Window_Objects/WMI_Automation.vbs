Dim objWMIService, objItems, objItem
Dim strComputer

strComputer = "." ' Local machine

' Connect to WMI service
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

' Get CPU information
Set objItems = objWMIService.ExecQuery("Select * from Win32_Processor")

' Display information
For Each objItem In objItems
    WScript.Echo "CPU: " & objItem.Name
Next

' Clean up
Set objItems = Nothing
Set objWMIService = Nothing

' Output :CPU: Intel(R) Core(TM) i3-6006U CPU @ 2.00GHz

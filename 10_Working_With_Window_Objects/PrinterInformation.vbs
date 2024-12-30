Dim objWMIService, objPrinter, colPrinters

' Connect to the WMI service
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

' Query for all installed printers
Set colPrinters = objWMIService.ExecQuery("Select * from Win32_Printer")

' Loop through the printers and display their names
For Each objPrinter In colPrinters
    WScript.Echo "Printer Name: " & objPrinter.Name
Next

' Clean up
Set colPrinters = Nothing
Set objWMIService = Nothing

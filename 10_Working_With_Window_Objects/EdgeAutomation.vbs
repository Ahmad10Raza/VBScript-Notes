Dim objIE

' Create an instance of Internet Explorer
Set objIE = CreateObject("InternetExplorer.Application")

' Make IE visible
objIE.Visible = True

' Navigate to a URL
objIE.Navigate "https://www.google.com"

' Wait for the page to load
Do While objIE.Busy Or objIE.readyState <> 4
    WScript.Sleep 100
Loop

' Close IE
objIE.Quit

' Clean up
Set objIE = Nothing

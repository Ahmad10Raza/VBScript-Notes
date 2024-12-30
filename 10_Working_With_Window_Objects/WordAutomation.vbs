Dim objWord, objDoc

' Create a Word application object
Set objWord = CreateObject("Word.Application")

' Make Word visible
objWord.Visible = True

' Create a new document
Set objDoc = objWord.Documents.Add()

' Write text to the document
objDoc.Content.Text = "Hello from VBScript in Word"

' Save the document
objDoc.SaveAs "D:\MSLTech\VBScript-Notes\10_Working_With_Window_Objects\example.docx"

' Close the document
objDoc.Close

' Quit Word
objWord.Quit

' Clean up
Set objDoc = Nothing
Set objWord = Nothing

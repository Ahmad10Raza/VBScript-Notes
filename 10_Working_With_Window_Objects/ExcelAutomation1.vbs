Dim objExcel, objWorkbook

' Create an Excel application object
Set objExcel = CreateObject("Excel.Application")

' Make Excel visible
objExcel.Visible = True

' Create a new workbook
Set objWorkbook = objExcel.Workbooks.Add()

' Write data to the first cell
objWorkbook.Sheets(1).Cells(1, 1).Value = "Hello from VBScript"

' Save the workbook
objWorkbook.SaveAs "D:\MSLTech\VBScript-Notes\10_Working_With_Window_Objects\example.xlsx"

' Close the workbook
objWorkbook.Close

' Quit Excel
objExcel.Quit

' Clean up
Set objWorkbook = Nothing
Set objExcel = Nothing

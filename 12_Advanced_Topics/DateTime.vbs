' Example Script: Date Manipulation and Formatting
Dim today, nextWeek, daysDifference

' Get current date
today = Date
MsgBox "Today's Date: " & today

' Add 7 days to the current date
nextWeek = DateAdd("d", 7, today)
MsgBox "Date Next Week: " & nextWeek

' Calculate days until a specific date
Dim targetDate
targetDate = #12/31/2024#
daysDifference = DateDiff("d", today, targetDate)
MsgBox "Days until " & targetDate & ": " & daysDifference

' Extract components of the current date
MsgBox "Year: " & Year(today) & ", Month: " & Month(today) & ", Day: " & Day(today)

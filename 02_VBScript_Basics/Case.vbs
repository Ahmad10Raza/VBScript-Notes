Dim day
day = Weekday(Now)
Select Case day
    Case 1: MsgBox "Sunday"
    Case 2: MsgBox "Monday"
    Case Else: MsgBox "Another Day"
End Select

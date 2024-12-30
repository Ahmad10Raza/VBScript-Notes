Dim num1, num2, result

num1 = 10
num2 = 5

' Using And
If (num1 > num2) And (num2 > 0) Then
    result = "num1 is greater than num2 and num2 is positive"
End If

' Using Or
If (num1 > 5) Or (num2 < 0) Then
    result = result & vbCrLf & "One of the conditions is true"
End If

' Using Not
If Not (num1 < num2) Then
    result = result & vbCrLf & "num1 is not less than num2"
End If

' Display result
MsgBox result

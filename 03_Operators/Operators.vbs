Dim num1, num2, addition, subtraction, multiplication, division, intDivision, modulus, exponentiation

num1 = 10
num2 = 3

addition = num1 + num2  ' 10 + 3 = 13
subtraction = num1 - num2  ' 10 - 3 = 7
multiplication = num1 * num2  ' 10 * 3 = 30
division = num1 / num2  ' 10 / 3 = 3.333...
intDivision = num1 \ num2  ' 10 \ 3 = 3 (integer part)
modulus = num1 Mod num2  ' 10 Mod 3 = 1 (remainder)
exponentiation = 2 ^ 4  ' 2^4 = 16

' Display results
MsgBox "Addition: " & addition & vbCrLf & _
       "Subtraction: " & subtraction & vbCrLf & _
       "Multiplication: " & multiplication & vbCrLf & _
       "Division: " & division & vbCrLf & _
       "Integer Division: " & intDivision & vbCrLf & _
       "Modulus: " & modulus & vbCrLf & _
       "Exponentiation: " & exponentiation

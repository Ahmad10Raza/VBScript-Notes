Dim matrix(2, 2)  ' 3x3 matrix

' Initializing values
matrix(0, 0) = 1
matrix(0, 1) = 2
matrix(0, 2) = 3
matrix(1, 0) = 4
matrix(1, 1) = 5
matrix(1, 2) = 6
matrix(2, 0) = 7
matrix(2, 1) = 8
matrix(2, 2) = 9

' Iterating through the matrix and displaying values
Dim row, col
For row = 0 To 2
    For col = 0 To 2
        MsgBox "Matrix[" & row & "," & col & "]: " & matrix(row, col)
    Next
Next

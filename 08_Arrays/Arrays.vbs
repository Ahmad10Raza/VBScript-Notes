' Static Array Example
Dim colors(2)
colors(0) = "Red"
colors(1) = "Green"
colors(2) = "Blue"
MsgBox "Color 1: " & colors(0) ' Output: Red

' Dynamic Array Example
Dim numbers()
ReDim numbers(1)
numbers(0) = 10
numbers(1) = 20
ReDim Preserve numbers(3)
numbers(2) = 30
numbers(3) = 40
MsgBox "Number 2: " & numbers(2) ' Output: 30

' Multidimensional Array Example
Dim grid(2, 2)
grid(0, 0) = 1
grid(0, 1) = 2
grid(1, 0) = 3
grid(1, 1) = 4
MsgBox "Grid[1,1]: " & grid(1, 1) ' Output: 4

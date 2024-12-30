' VBScript: Looping Examples

' ------------------------------
' 1. For...Next Loop Example
' ------------------------------
Dim i
For i = 1 To 5
    MsgBox "For...Next Loop - Iteration " & i
Next

' ------------------------------
' 2. For Each...Next Loop Example
' ------------------------------
Dim fruits, fruit
fruits = Array("Apple", "Banana", "Cherry")

For Each fruit In fruits
    MsgBox "For Each...Next Loop - Fruit: " & fruit
Next

' ------------------------------
' 3. Do While...Loop Example
' ------------------------------
Dim j
j = 1
Do While j <= 5
    MsgBox "Do While...Loop - j = " & j
    j = j + 1
Loop

' ------------------------------
' 4. Do Until...Loop Example
' ------------------------------
Dim k
k = 1
Do Until k > 5
    MsgBox "Do Until...Loop - k = " & k
    k = k + 1
Loop

' ------------------------------
' 5. While...Wend Loop Example
' ------------------------------
Dim m
m = 1
While m <= 5
    MsgBox "While...Wend Loop - m = " & m
    m = m + 1
Wend

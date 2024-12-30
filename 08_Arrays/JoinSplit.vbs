' Join example
Dim fruits(2)
fruits(0) = "Apple"
fruits(1) = "Banana"
fruits(2) = "Cherry"

Dim fruitString
fruitString = Join(fruits, ", ")  ' Join array elements with a comma and space
MsgBox fruitString  ' Output: Apple, Banana, Cherry

' Split example
Dim splitFruits
splitFruits = Split(fruitString, ", ")  ' Split the string back into an array

' Displaying the split array
For i = LBound(splitFruits) To UBound(splitFruits)
    MsgBox "Split Fruits(" & i & "): " & splitFruits(i)
Next

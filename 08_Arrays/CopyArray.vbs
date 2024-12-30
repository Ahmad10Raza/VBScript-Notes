Dim originalArray(2)
originalArray(0) = "Apple"
originalArray(1) = "Banana"
originalArray(2) = "Cherry"

' Declare a new array and copy elements
Dim copiedArray(2)
For i = LBound(originalArray) To UBound(originalArray)
    copiedArray(i) = originalArray(i)
Next

' Display copied array elements
For i = LBound(copiedArray) To UBound(copiedArray)
    MsgBox "Copied Array(" & i & "): " & copiedArray(i)
Next

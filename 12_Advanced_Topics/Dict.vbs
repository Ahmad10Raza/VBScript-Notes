' Dictionary Object Demonstration
Dim dict
Set dict = CreateObject("Scripting.Dictionary")

' Section 1: Adding Items
dict.Add "name", "Alice"
dict.Add "age", 25
dict.Add "country", "USA"

MsgBox "Added items to the dictionary."

' Section 2: Accessing Values
MsgBox "Name: " & dict("name")
MsgBox "Age: " & dict("age")
MsgBox "Country: " & dict("country")

' Section 3: Checking Key Existence
If dict.Exists("age") Then
    MsgBox "Key 'age' exists with value: " & dict("age")
Else
    MsgBox "Key 'age' does not exist."
End If

If dict.Exists("city") Then
    MsgBox "Key 'city' exists."
Else
    MsgBox "Key 'city' does not exist."
End If

' Section 4: Removing an Item
dict.Remove "age"
If Not dict.Exists("age") Then
    MsgBox "Key 'age' removed successfully."
End If

' Section 5: Iterating Through Dictionary
Dim key
For Each key In dict.Keys
    MsgBox "Key: " & key & ", Value: " & dict(key)
Next

' Section 6: Getting the Count
MsgBox "Number of items in the dictionary: " & dict.Count

' Section 7: Clearing the Dictionary
dict.RemoveAll
MsgBox "All items removed. Dictionary count: " & dict.Count

' Clean up
Set dict = Nothing
MsgBox "Dictionary operations complete."

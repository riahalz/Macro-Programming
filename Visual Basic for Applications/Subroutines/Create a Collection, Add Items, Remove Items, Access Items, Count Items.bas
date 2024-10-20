' Create a collection, add items, access items, remove items, count items

Sub CollectionExample()

‘ Create collection
Dim fruits As New Collection

' Add items to collection
fruits.Add "Apple"
fruits.Add "Banana"
fruits.Add “Cherry"

' Access collection items
Debug.Print fruits(2)  ' Output: Banana

' Remove an item
fruits.Remove (1)

' Count items
Debug.Print fruits.Count ' Output: 2’

End Sub

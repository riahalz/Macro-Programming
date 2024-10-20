' Property examples - Read-only, Write-only, Read-Write

Sub PropertyExamples()

    Dim ws As Worksheet
    Dim rng As Range

    ' Set the active worksheet
    Set ws = ActiveSheet

    ' Read-only property
    Debug.Print "Current worksheet name: " & ws.Name

    ' Write properties
    rng.Value = Array("Product", "Price", "Apple", 0.5)
    rng.Font.Bold = True
    rng.Interior.Color = RGB(255, 255, 0) ' Yellow

    ' Read-write property
    rng.Cells(2, 2).Value = rng.Cells(2, 2).Value * 2
    Debug.Print "New price: " & rng.Cells(2, 2).Value

End Sub

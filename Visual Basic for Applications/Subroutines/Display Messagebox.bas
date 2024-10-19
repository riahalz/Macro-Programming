Sub BasicExample()
    ' Declare variables
    Dim cellValue As Integer
    Dim message As String

    ' Assign a value to a cell
    Range("A1").Value = 10

    ' Read the value from the cell
    cellValue = Range("A1").Value

    ' Use an If–Then statement
    If cellValue > 5 Then
        message = "The value is greater than 5"
    Else
        message = "The value is 5 or less"
    End If

    ' Display the message in a message box
    MsgBox message, vbInformation, "Result"

    ' Use a For loop to fill a range of cells
    For i = 1 To 5
        Cells(i, 2).Value = i * 2
    Next i

    ' Change cell color
    Range("B1:B5").Interior.Color = RGB(255, 255, 0)
End Sub

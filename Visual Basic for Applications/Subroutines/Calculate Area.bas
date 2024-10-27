Sub Calculate Area()

    Dim length As Variant
    Dim width As Variant
    Dim area As Double

    ' Ask user to input length and width
    length = Application.InputBox("Input length:", Type:=1) ' Type:=1 allows only numerical input
    If length = False Then Exit Sub ' User cancelled

    width = Application.InputBox("Input width:", Type:=1)
    If width = False Then Exit Sub ' User cancelled

    ' Check validity of inputs
    If Not IsNumeric(length) Or Not IsNumeric(width) Then
        MsgBox "Please enter valid numeric values for length and width.", vbExclamation
        Exit Sub
    End If

    ' Convert length and width to Double
    Dim lengthDouble As Double
    Dim widthDouble As Double

    lengthDouble = CDbl(length)
    widthDouble = CDbl(width)

    ' Calculate area using custom function
    area = CalculateArea(lengthDouble, widthDouble)

    With ThisWorkbook.Worksheets("Sheet1")
        ' Display values in cells
        .Range("A1").Value = "Length"
        .Range("B1").Value = "Width"
        .Range("A2").Value = lengthDouble
        .Range("B2").Value = widthDouble
    End With

    ' Call recorded cell formatting macro
    Call A1B1BoldCenter

    ' Display calculated area in message box
    MsgBox "The calculated area is " & area, vbInformation, "Area Calculation"

End Sub

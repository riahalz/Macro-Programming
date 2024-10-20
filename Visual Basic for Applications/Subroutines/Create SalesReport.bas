' Create SalesReport with data

Sub CreateSalesReport()
    ' Declare variables
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range

    ' Create a new workbook
    Set wb = Application.Workbooks.Add

    ' Set the active sheet and rename it
    Set ws = wb.ActiveSheet
    ws.Name = "Sales Report"

    ' Add headers
    ws.Range("A1").Value = "Product"
    ws.Range("B1").Value = "Quantity"
    ws.Range("C1").Value = "Price"
    ws.Range("D1").Value = "Total"

    ' Add some sample data
    ws.Range("A2:A4").Value = Application.Transpose(Array("Apples", "Bananas", "Oranges"))
    ws.Range("B2:B4").Value = Application.Transpose(Array(100, 150, 75))
    ws.Range("C2:C4").Value = Application.Transpose(Array(0.5, 0.3, 0.6))

    ' Calculate totals
    Set rng = ws.Range("D2:D4")
    rng.FormulaR1C1 = "=RC[-2]*RC[-1]"

    ' Format headers
    ws.Range("A1:D1").Font.Bold = True

    ' Autofit columns
    ws.Columns("A:D").AutoFit

    ' Save the workbook
    wb.SaveAs FileName:="SalesReport.xlsx"

    ' Close the workbook
    wb.Close SaveChanges:=True

    MsgBox "Sales report has been created and saved as 'SalesReport.xlsx'", vbInformation
End Sub

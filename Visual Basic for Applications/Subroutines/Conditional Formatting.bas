Sub ConditionalFormatting()
Dim rng As Range
Set rng = ActiveSheet.Range("A1:A10")
rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=50"
With rng.FormatConditions(rng.FormatConditions.Count)
.Interior.Color = RGB(255, 0, 0) 'Red
End With
End Sub

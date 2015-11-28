Attribute VB_Name = "Module4"
Sub redoFormatting()
Attribute redoFormatting.VB_ProcData.VB_Invoke_Func = " \n14"
Range("C3").Activate
Cells.FormatConditions.Delete
Range("C3:Q17").FormatConditions.Add Type:=xlExpression, Formula1:="=BE21=5"
Cells.FormatConditions(1).Interior.color = RGB(0, 0, 0)
Range("BD38:BT54").FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=1"
Cells.FormatConditions(2).Interior.color = 13551615
Range("BD38:BT54").FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=-1"
Cells.FormatConditions(3).Interior.color = 13561798
End Sub

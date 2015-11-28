Attribute VB_Name = "Module1"
Sub nextCell(currentCell As Range, topORleft As Boolean, alpha As Double, offsetR As Integer, offsetC As Integer)
Dim height As Double, length As Double
Dim theta As Double
height = 0.2
length = 0.2
theta = Application.WorksheetFunction.Pi() * 25 / 180
With currentCell.Interior
   .Pattern = xlSolid
   .PatternColorIndex = xlAutomatic
   .ThemeColor = xlThemeColorLight1
   .TintAndShade = 0
   .PatternTintAndShade = 0
End With
If height / length > Tan(theta) Then
   If topORleft Then
      Set currentCell = currentCell.Offset(0, offsetC)
      topORleft = Not topORleft
      alpha = (length - alpha) * Tan(theta)
   ElseIf alpha + length * Tan(theta) < height Then
      Set currentCell = currentCell.Offset(0, offsetC)
      alpha = alpha + length * Tan(theta)
   Else
      Set currentCell = currentCell.Offset(offsetR, 0)
      topORleft = Not topORleft
      alpha = (height - alpha) / Tan(theta)
   End If
Else
   If Not topORleft Then
      Set currentCell = currentCell.Offset(offsetR, 0)
      topORleft = Not topORleft
      alpha = (height - alpha) / Tan(theta)
   ElseIf alpha + height / Tan(theta) < length Then
      Set currentCell = currentCell.Offset(offsetR, 0)
      alpha = alpha + height / Tan(theta)
   Else
      Set currentCell = currentCell.Offset(0, offsetC)
      topORleft = Not topORleft
      alpha = (length - alpha) * Tan(theta)
   End If
End If
End Sub

Sub drawLine()
Cells.Clear
Dim currentCell_1 As Range, currentCell_2 As Range
Set currentCell_1 = Range("Y50")
Set currentCell_2 = Range("Y50")
Dim topORleft_1 As Boolean, topOrleft_2 As Boolean
topORleft_1 = True
topOrleft_2 = False
Dim alpha_1 As Double, alpha_2 As Double
alpha_1 = 0
alpha_2 = 0
For i = 1 To 100
Call nextCell(currentCell_1, topORleft_1, alpha_1, 1, 1)
Call nextCell(currentCell_2, topOrleft_2, alpha_2, -1, 1)
Next i
End Sub

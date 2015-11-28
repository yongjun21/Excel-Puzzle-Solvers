Attribute VB_Name = "Module1"
Sub main()
Range("B2:J10").Font.Bold = False
Range("B2:J10").Font.Color = RGB(0, 0, 0)
Range("B15:J95").Value = 1
For row = 0 To 8
For col = 0 To 8
If Range("B2").Offset(row, col).Value <> vbNullString Then
    Range("B2").Offset(row, col).Font.Bold = True
    Range("B2").Offset(row, col).Font.Color = RGB(255, 0, 0)
    Call WriteAns(row, col, Range("B2").Offset(row, col).Value)
End If
Next col
Next row
If Narrow() = 0 Then Call Trial
End Sub


Function Narrow() As Integer
Do
Loop Until Not Eliminate()

If WorksheetFunction.Count(Range("B2:J10")) = 81 Then
    Narrow = 1
Else
   For i = 0 To 80
      If Range("B2").Offset(i \ 9, i Mod 9).Value = vbNullString And _
         WorksheetFunction.Sum(Range("A15:I15").Offset(i, 1)) = 0 Then
         Narrow = -1
         Exit Function
      End If
   Next i
   
   For j = 1 To 9
   For k = 0 To 8
      If WorksheetFunction.CountIf(Range("B2:J2").Offset(k, 0), j) = 0 And _
         SumRow(k, j) = 0 Then
         Narrow = -1
         Exit Function
      End If
      
      If WorksheetFunction.CountIf(Range("B2:B10").Offset(0, k), j) = 0 And _
         SumCol(k, j) = 0 Then
         Narrow = -1
         Exit Function
      End If
   
      If WorksheetFunction.CountIf(Range("B2:D4").Offset(3 * (k \ 3), 3 * (k Mod 3)), j) = 0 And _
         SumBox(k, j) = 0 Then
         Narrow = -1
         Exit Function
      End If
   Next k
   Next j
End If
End Function


Function Eliminate() As Boolean
For j = 1 To 9
For k = 0 To 8
   If SumRow(k, j) = 1 Then
      For ii = 0 To 8
         If Range("A15").Offset(k * 9 + ii, j).Value = 1 Then
            Call WriteAns(k, ii, j)
            Eliminate = True
            Exit For
         End If
      Next ii
   End If
   
   If SumCol(k, j) = 1 Then
      For ii = 0 To 8
         If Range("A15").Offset(ii * 9 + k, j).Value = 1 Then
            Call WriteAns(ii, k, j)
            Eliminate = True
            Exit For
         End If
      Next ii
   End If

   If SumBox(k, j) = 1 Then
      For ii = 0 To 8
         If Range("A15").Offset(3 * BoxStep(k) + BoxStep(ii), j).Value = 1 Then
            Call WriteAns(3 * (k \ 3) + (ii \ 3), 3 * (k Mod 3) + (ii Mod 3), j)
            Eliminate = True
            Exit For
         End If
      Next ii
   End If
Next k
Next j

For i = 0 To 80
   If WorksheetFunction.Sum(Range("A15:I15").Offset(i, 1)) = 1 Then
   For j = 1 To 9
      If Range("A15").Offset(i, j).Value = 1 Then
         Call WriteAns(i \ 9, i Mod 9, j)
         Eliminate = True
         Exit For
      End If
   Next j
   End If
Next i
End Function


Function Trial() As Integer
Dim Result As Integer
Dim State1 As Variant, State2 As Variant
Dim random As Integer

State1 = Range("B2:J10").Value
State2 = Range("B15:J95").Value

random = Int(81 * Rnd)
For i = 0 To 80
If WorksheetFunction.Sum(Range("A15:I15").Offset((i + random) Mod 81, 1)) > 0 Then Exit For
Next i
i = (i + random) Mod 81

For j = 1 To 9
   If Range("A15").Offset(i, j).Value = 1 Then
      Call WriteAns(i \ 9, i Mod 9, j)
      Trial = Narrow()
   
      If Trial = 1 Then
         Exit For
      ElseIf Trial = 0 Then
         Trial = Trial()
         If Trial = 1 Then Exit For
      End If
         
      Range("B2:J10").Value = State1
      Range("B15:J95").Value = State2
   End If
Next j
End Function


Sub WriteAns(row, col, Ans)
Range("B2").Offset(row, col).Value = Ans
For ii = 0 To 8
   Range("A15").Offset(row * 9 + ii, Ans).Value = 0
   Range("A15").Offset(ii * 9 + col, Ans).Value = 0
   Range("A15").Offset(3 * BoxStep(3 * (row \ 3) + (col \ 3)) + BoxStep(ii), Ans).Value = 0
   Range("A15").Offset(row * 9 + col, ii + 1).Value = 0
Next ii
End Sub

Function SumRow(k, j) As Integer
SumRow = WorksheetFunction.Sum(Range("A15:A23").Offset(k * 9, j))
End Function

Function SumCol(k, j) As Integer
For ii = 0 To 8
SumCol = SumCol + Range("A15").Offset(ii * 9 + k, j).Value
Next ii
End Function

Function SumBox(k, j) As Integer
For ii = 0 To 8
SumBox = SumBox + Range("A15").Offset(3 * BoxStep(k) + BoxStep(ii), j).Value
Next ii
End Function

Function BoxStep(n) As Integer
Dim n1 As Integer, n2 As Integer
n1 = n \ 3
n2 = n Mod 3
BoxStep = n1 * 9 + n2
End Function

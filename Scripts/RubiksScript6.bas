Attribute VB_Name = "Module6"
Sub solveCube()
Sheets("Moves").Range("A3:N1002").Copy destination:=Sheets("Moves").Range("D3")
Sheets("Moves").Range("A3:B1002").Clear

Dim firstMove As Range, nextMove As Range
Set firstMove = Sheets("Moves").Range("A3")
Dim savestate As Variant
savestate = Sheets("Mapping").Range("E2:F55").Value

Call step_0(firstMove)
Set nextMove = firstMove

Call step_1(nextMove)
Call step_2(nextMove)
Call step_3(nextMove)
Call step_4(nextMove)
Call step_5(nextMove)
Call step_6(nextMove)
Call step_7(nextMove)
Call extra_step(firstMove, nextMove.Row - firstMove.Row - 1)

Sheets("Mapping").Range("E2:F55").Value = savestate
executeMoves (Sheets("Moves").Range("A3"))
End Sub


Sub step_0(firstMove As Range)
For k = 0 To 3
   If Sheets("Mapping").Range("F6").Value = 0 Then Exit For
   Call rotate(2, True, 3, firstMove)
Next k
If k = 4 Then
For k = 0 To 3
   If Sheets("Mapping").Range("F6").Value = 0 Then Exit For
   Call rotate(4, True, 3, firstMove)
Next k
End If
For k = 0 To 3
   If Sheets("Mapping").Range("F24").Value = 2 Then Exit For
   Call rotate(0, True, 3, firstMove)
Next k

Call simplify2(Sheets("Moves").Range("A3"))
Set firstMove = Sheets("Moves").Range("A3")
Do While firstMove <> vbNullString
Set firstMove = firstMove.Offset(1, 0)
Loop
End Sub


Sub step_1(nextMove As Range)
Dim changeSide As Integer
Dim colorArray() As Variant
colorArray = Array(2, 1, 5, 4)

For Each targetColor In colorArray
For changeSide = 0 To 3
   If Sheets("Mapping").Range("F41").Value = 3 And Sheets("Mapping").Range("F21").Value = targetColor Then
      For k = 1 To changeSide
         Call rotate(3, True, 1, nextMove)
      Next k
      Call rotate(2, False, 1, nextMove)
      Call rotate(3, False, 1, nextMove)
      Exit For
   ElseIf Sheets("Mapping").Range("F21").Value = 3 And Sheets("Mapping").Range("F41").Value = targetColor Then
      For k = 1 To changeSide
         Call rotate(3, True, 1, nextMove)
      Next k
      Call rotate(3, False, 1, nextMove)
      Call rotate(4, True, 1, nextMove)
      Exit For
   ElseIf Sheets("Mapping").Range("F3").Value = 3 And Sheets("Mapping").Range("F23").Value = targetColor Then
      For k = 1 To changeSide
         Call rotate(3, True, 1, nextMove)
      Next k
      Call rotate(2, False, 1, nextMove)
      Call rotate(2, False, 1, nextMove)
      Call rotate(3, False, 1, nextMove)
      Exit For
   ElseIf Sheets("Mapping").Range("F23").Value = 3 And Sheets("Mapping").Range("F3").Value = targetColor Then
      For k = 1 To changeSide
         Call rotate(3, True, 1, nextMove)
      Next k
      Call rotate(2, False, 1, nextMove)
      Call rotate(3, False, 1, nextMove)
      Call rotate(4, True, 1, nextMove)
      Exit For
   ElseIf Sheets("Mapping").Range("F30").Value = 3 And Sheets("Mapping").Range("F25").Value = targetColor Then
      Call rotate(2, True, 1, nextMove)
      For k = 1 To changeSide
         Call rotate(3, True, 1, nextMove)
      Next k
      Call rotate(2, False, 1, nextMove)
      Call rotate(3, False, 1, nextMove)
      Exit For
   ElseIf Sheets("Mapping").Range("F25").Value = 3 And Sheets("Mapping").Range("F30").Value = targetColor Then
      Call rotate(2, True, 1, nextMove)
      For k = 1 To changeSide
         Call rotate(3, True, 1, nextMove)
      Next k
      Call rotate(3, False, 1, nextMove)
      Call rotate(4, True, 1, nextMove)
      Exit For
   End If
   Call rotate(0, True, 3, nextMove)
Next changeSide
Next
End Sub


Sub step_2(nextMove As Range)
Dim changeSide As Integer
Dim colorArray() As Variant
colorArray = Array(2, 1, 5, 4)

For Each targetColor In colorArray
For changeSide = 0 To 3
   If Sheets("Mapping").Range("F20").Value = 3 And Sheets("Mapping").Range("F2").Value = targetColor Then
      For k = 1 To changeSide
         Call rotate(3, True, 1, nextMove)
      Next k
      Call rotate(2, True, 1, nextMove)
      Call rotate(0, True, 1, nextMove)
      Call rotate(2, False, 1, nextMove)
      Call rotate(3, False, 1, nextMove)
      Exit For
   ElseIf Sheets("Mapping").Range("F38").Value = 3 And Sheets("Mapping").Range("F20").Value = targetColor Then
      For k = 1 To changeSide
         Call rotate(3, True, 1, nextMove)
      Next k
      Call rotate(4, False, 1, nextMove)
      Call rotate(0, False, 1, nextMove)
      Call rotate(4, True, 1, nextMove)
      Call rotate(3, False, 1, nextMove)
      Exit For
   ElseIf Sheets("Mapping").Range("F2").Value = 3 And Sheets("Mapping").Range("F38").Value = targetColor Then
      For k = 1 To changeSide
         Call rotate(3, True, 1, nextMove)
      Next k
      Call rotate(2, True, 1, nextMove)
      Call rotate(0, False, 1, nextMove)
      Call rotate(2, False, 1, nextMove)
      Call rotate(0, True, 1, nextMove)
      Call rotate(4, False, 1, nextMove)
      Call rotate(0, True, 1, nextMove)
      Call rotate(4, True, 1, nextMove)
      Call rotate(3, False, 1, nextMove)
      Exit For
   ElseIf Sheets("Mapping").Range("F22").Value = 3 And Sheets("Mapping").Range("F44").Value = targetColor Then
      Call rotate(2, True, 1, nextMove)
      Call rotate(0, True, 1, nextMove)
      Call rotate(2, False, 1, nextMove)
      Call rotate(0, False, 1, nextMove)
      For k = 1 To changeSide
         Call rotate(3, True, 1, nextMove)
      Next k
      Call rotate(2, True, 1, nextMove)
      Call rotate(0, True, 1, nextMove)
      Call rotate(2, False, 1, nextMove)
      Call rotate(3, False, 1, nextMove)
      Exit For
   ElseIf Sheets("Mapping").Range("F44").Value = 3 And Sheets("Mapping").Range("F29").Value = targetColor Then
      Call rotate(4, False, 1, nextMove)
      Call rotate(0, False, 1, nextMove)
      Call rotate(4, True, 1, nextMove)
      Call rotate(0, True, 1, nextMove)
      For k = 1 To changeSide
         Call rotate(3, True, 1, nextMove)
      Next k
      Call rotate(4, False, 1, nextMove)
      Call rotate(0, False, 1, nextMove)
      Call rotate(4, True, 1, nextMove)
      Call rotate(3, False, 1, nextMove)
      Exit For
   ElseIf Sheets("Mapping").Range("F29").Value = 3 And Sheets("Mapping").Range("F22").Value = targetColor Then
      Call rotate(2, True, 1, nextMove)
      Call rotate(0, False, 1, nextMove)
      Call rotate(2, False, 1, nextMove)
      For k = 1 To changeSide
         Call rotate(3, True, 1, nextMove)
      Next k
      Call rotate(2, True, 1, nextMove)
      Call rotate(0, True, 1, nextMove)
      Call rotate(2, False, 1, nextMove)
      Call rotate(3, False, 1, nextMove)
      Exit For
   End If
   Call rotate(0, True, 3, nextMove)
Next changeSide
Next

For changeSide = 0 To 3
   If Sheets("Mapping").Range("F24").Value = 2 And Sheets("Mapping").Range("F42").Value = 4 Then
      For k = 1 To changeSide
         Call rotate(3, True, 1, nextMove)
      Next k
      Exit For
   End If
   Call rotate(0, True, 3, nextMove)
Next changeSide
End Sub


Sub step_3(nextMove As Range)
Dim changeSide As Integer
Dim colorArray1() As Variant, colorArray2() As Variant
colorArray1 = Array(2, 1, 5, 4)
colorArray2 = Array(4, 2, 1, 5)

For n = 0 To 3
For changeSide = 0 To 3
   If Sheets("Mapping").Range("F9").Value = colorArray1(n) And Sheets("Mapping").Range("F50").Value = colorArray2(n) Then
      For k = 1 To changeSide
         Call rotate(0, False, 3, nextMove)
         Call rotate(0, True, 1, nextMove)
      Next k
      Call rotate(2, True, 1, nextMove)
      Call rotate(0, False, 1, nextMove)
      Call rotate(2, False, 1, nextMove)
      Call rotate(0, False, 1, nextMove)
      Call rotate(4, False, 1, nextMove)
      Call rotate(0, True, 1, nextMove)
      Call rotate(4, True, 1, nextMove)
      Call rotate(0, True, 3, nextMove)
      Exit For
   ElseIf Sheets("Mapping").Range("F12").Value = colorArray1(n) And Sheets("Mapping").Range("F7").Value = colorArray2(n) Then
      For k = 1 To changeSide
         Call rotate(0, False, 3, nextMove)
         Call rotate(0, True, 1, nextMove)
      Next k
      Call rotate(4, False, 1, nextMove)
      Call rotate(0, True, 1, nextMove)
      Call rotate(4, True, 1, nextMove)
      Call rotate(0, True, 1, nextMove)
      Call rotate(2, True, 1, nextMove)
      Call rotate(0, False, 1, nextMove)
      Call rotate(2, False, 1, nextMove)
      Call rotate(0, True, 3, nextMove)
      Exit For
   ElseIf Sheets("Mapping").Range("F21").Value = colorArray1(n) And Sheets("Mapping").Range("F41").Value = colorArray2(n) Then
      Call rotate(4, False, 1, nextMove)
      Call rotate(0, False, 1, nextMove)
      Call rotate(4, True, 1, nextMove)
      Call rotate(0, True, 1, nextMove)
      Call rotate(2, True, 1, nextMove)
      Call rotate(0, True, 1, nextMove)
      Call rotate(2, False, 1, nextMove)
      For k = 1 To changeSide
         Call rotate(0, False, 3, nextMove)
         Call rotate(0, True, 1, nextMove)
      Next k
      Call rotate(2, True, 1, nextMove)
      Call rotate(0, False, 1, nextMove)
      Call rotate(2, False, 1, nextMove)
      Call rotate(0, False, 1, nextMove)
      Call rotate(4, False, 1, nextMove)
      Call rotate(0, True, 1, nextMove)
      Call rotate(4, True, 1, nextMove)
      Call rotate(0, True, 3, nextMove)
      Exit For
   ElseIf Sheets("Mapping").Range("F41").Value = colorArray1(n) And Sheets("Mapping").Range("F21").Value = colorArray2(n) Then
      Call rotate(2, True, 1, nextMove)
      Call rotate(0, False, 1, nextMove)
      Call rotate(2, False, 1, nextMove)
      Call rotate(0, False, 1, nextMove)
      Call rotate(4, False, 1, nextMove)
      Call rotate(0, True, 1, nextMove)
      Call rotate(4, True, 1, nextMove)
      Call rotate(0, False, 1, nextMove)
      For k = 1 To changeSide
         Call rotate(0, False, 3, nextMove)
         Call rotate(0, True, 1, nextMove)
      Next k
      Call rotate(2, True, 1, nextMove)
      Call rotate(0, False, 1, nextMove)
      Call rotate(2, False, 1, nextMove)
      Call rotate(0, False, 1, nextMove)
      Call rotate(4, False, 1, nextMove)
      Call rotate(0, True, 1, nextMove)
      Call rotate(4, True, 1, nextMove)
      Call rotate(0, True, 3, nextMove)
      Exit For
   End If
   Call rotate(0, True, 3, nextMove)
Next changeSide
Next n
End Sub


Sub step_4(nextMove As Range)
For k = 0 To 3
   If Sheets("Mapping").Range("F3").Value > 0 And Sheets("Mapping").Range("F9").Value > 0 Then
      Call rotate(2, True, 1, nextMove)
      Call rotate(1, True, 1, nextMove)
      Call rotate(0, True, 1, nextMove)
      Call rotate(1, False, 1, nextMove)
      Call rotate(0, False, 1, nextMove)
      Call rotate(2, False, 1, nextMove)
   ElseIf Sheets("Mapping").Range("F3").Value > 0 And Sheets("Mapping").Range("F7").Value > 0 Then
      Call rotate(2, True, 1, nextMove)
      Call rotate(1, True, 1, nextMove)
      Call rotate(0, False, 1, nextMove)
      Call rotate(1, False, 1, nextMove)
      Call rotate(0, True, 1, nextMove)
      Call rotate(2, False, 1, nextMove)
      Call rotate(4, False, 1, nextMove)
      Call rotate(0, True, 1, nextMove)
      Call rotate(4, True, 1, nextMove)
   End If
   Call rotate(0, True, 3, nextMove)
Next k
End Sub


Sub step_5(nextMove As Range)
Dim colorArray1() As Variant, colorArray2() As Variant, colorArray3() As Variant
colorArray1 = Array(2, 1, 5, 4)
colorArray2 = Array(4, 2, 1, 5)
colorArray3 = Array(5, 4, 2, 1)

For n = 0 To 3
For k = 0 To 3
   If Sheets("Mapping").Range("F23").Value = colorArray1(n) Then Exit For
   Call rotate(0, True, 3, nextMove)
Next k
If Sheets("Mapping").Range("F39").Value = colorArray2(n) Then
   If Sheets("Mapping").Range("F50").Value = colorArray3(n) Then Exit For
   Call rotate(4, True, 1, nextMove)
   Call rotate(0, True, 1, nextMove)
   Call rotate(4, False, 1, nextMove)
   Call rotate(0, True, 1, nextMove)
   Call rotate(4, True, 1, nextMove)
   Call rotate(0, True, 1, nextMove)
   Call rotate(0, True, 1, nextMove)
   Call rotate(4, False, 1, nextMove)
   Exit For
ElseIf Sheets("Mapping").Range("F50").Value = colorArray3(n) Then
   Call rotate(4, True, 1, nextMove)
   Call rotate(0, True, 1, nextMove)
   Call rotate(4, False, 1, nextMove)
   Call rotate(0, True, 1, nextMove)
   Call rotate(4, True, 1, nextMove)
   Call rotate(0, True, 1, nextMove)
   Call rotate(0, True, 1, nextMove)
   Call rotate(4, False, 1, nextMove)
   Call rotate(0, False, 3, nextMove)
   Call rotate(4, True, 1, nextMove)
   Call rotate(0, True, 1, nextMove)
   Call rotate(4, False, 1, nextMove)
   Call rotate(0, True, 1, nextMove)
   Call rotate(4, True, 1, nextMove)
   Call rotate(0, True, 1, nextMove)
   Call rotate(0, True, 1, nextMove)
   Call rotate(4, False, 1, nextMove)
   Exit For
End If
Next n

For k = 0 To 3
   If Sheets("Mapping").Range("F24").Value = 2 Then Exit For
   Call rotate(0, True, 3, nextMove)
Next k
For k = 0 To 3
   If Sheets("Mapping").Range("F23").Value = 2 Then Exit For
   Call rotate(0, True, 1, nextMove)
Next k
End Sub


Sub step_6(nextMove As Range)
Dim colorArray1() As Variant, colorArray2() As Variant
colorArray1 = Array(2, 1, 5, 4)
colorArray2 = Array(4, 2, 1, 5)

Do
For n = 0 To 3
   If Sheets("Mapping").Range("F2").Value <> (colorArray1(n) + 3) Mod 6 And _
      Sheets("Mapping").Range("F20").Value <> (colorArray1(n) + 3) Mod 6 And _
      Sheets("Mapping").Range("F38").Value <> (colorArray1(n) + 3) Mod 6 And _
      Sheets("Mapping").Range("F2").Value <> (colorArray2(n) + 3) Mod 6 And _
      Sheets("Mapping").Range("F20").Value <> (colorArray2(n) + 3) Mod 6 And _
      Sheets("Mapping").Range("F38").Value <> (colorArray2(n) + 3) Mod 6 Then Exit Do
   Call rotate(0, True, 3, nextMove)
Next n
Call rotate(0, True, 1, nextMove)
Call rotate(2, True, 1, nextMove)
Call rotate(0, False, 1, nextMove)
Call rotate(5, False, 1, nextMove)
Call rotate(0, True, 1, nextMove)
Call rotate(2, False, 1, nextMove)
Call rotate(0, False, 1, nextMove)
Call rotate(5, True, 1, nextMove)
Loop

If Sheets("Mapping").Range("F4").Value <> colorArray1(n) And _
   Sheets("Mapping").Range("F11").Value <> colorArray1(n) And _
   Sheets("Mapping").Range("F26").Value <> colorArray1(n) And _
   Sheets("Mapping").Range("F4").Value <> colorArray2(n) And _
   Sheets("Mapping").Range("F11").Value <> colorArray2(n) And _
   Sheets("Mapping").Range("F26").Value <> colorArray2(n) Then
      Call rotate(0, True, 1, nextMove)
      Call rotate(2, True, 1, nextMove)
      Call rotate(0, False, 1, nextMove)
      Call rotate(5, False, 1, nextMove)
      Call rotate(0, True, 1, nextMove)
      Call rotate(2, False, 1, nextMove)
      Call rotate(0, False, 1, nextMove)
      Call rotate(5, True, 1, nextMove)
ElseIf Sheets("Mapping").Range("F8").Value <> colorArray1(n) And _
   Sheets("Mapping").Range("F40").Value <> colorArray1(n) And _
   Sheets("Mapping").Range("F47").Value <> colorArray1(n) And _
   Sheets("Mapping").Range("F8").Value <> colorArray2(n) And _
   Sheets("Mapping").Range("F40").Value <> colorArray2(n) And _
   Sheets("Mapping").Range("F47").Value <> colorArray2(n) Then
      Call rotate(0, False, 1, nextMove)
      Call rotate(4, False, 1, nextMove)
      Call rotate(0, True, 1, nextMove)
      Call rotate(1, True, 1, nextMove)
      Call rotate(0, False, 1, nextMove)
      Call rotate(4, True, 1, nextMove)
      Call rotate(0, True, 1, nextMove)
      Call rotate(1, False, 1, nextMove)
End If
End Sub


Sub step_7(nextMove As Range)
Dim cycleCount As Integer
Dim clockwise As Boolean

For k = 0 To 3
Do While Sheets("Mapping").Range("F2").Value > 0
   If cycleCount Mod 3 = 0 Then
      If Sheets("Mapping").Range("F20").Value = 0 Then clockwise = True Else clockwise = False
   End If
   If clockwise Then
      Call rotate(2, False, 1, nextMove)
      Call rotate(3, False, 1, nextMove)
      Call rotate(2, True, 1, nextMove)
      Call rotate(3, True, 1, nextMove)
      Call rotate(2, False, 1, nextMove)
      Call rotate(3, False, 1, nextMove)
      Call rotate(2, True, 1, nextMove)
      Call rotate(3, True, 1, nextMove)
   Else
      Call rotate(4, True, 1, nextMove)
      Call rotate(3, True, 1, nextMove)
      Call rotate(4, False, 1, nextMove)
      Call rotate(3, False, 1, nextMove)
      Call rotate(4, True, 1, nextMove)
      Call rotate(3, True, 1, nextMove)
      Call rotate(4, False, 1, nextMove)
      Call rotate(3, False, 1, nextMove)
   End If
   cycleCount = cycleCount + 1
Loop
Call rotate(0, True, 1, nextMove)
Next k
End Sub


Sub extra_step(firstMove As Range, lastR As Integer)
Dim advanceR As Integer
For advanceR = lastR To 0 Step -1
If firstMove.Offset(advanceR, 1).Value = 0 Then
   Call sendToBack(advanceR, lastR, firstMove)
   firstMove.Range("A1:B1").Offset(lastR, 0).ClearContents
   lastR = lastR - 1
End If
Next advanceR

Do
lastR0 = lastR

advanceR = lastR
Do While advanceR >= 0
If firstMove.Offset(advanceR + 1, 0).Value <> vbNullString Then
If firstMove.Offset(advanceR, 0).Value = firstMove.Offset(advanceR + 1, 0).Value And _
   firstMove.Offset(advanceR, 1).Value = -firstMove.Offset(advanceR + 1, 1).Value Then
      If advanceR + 1 < lastR Then Range(firstMove.Offset(advanceR + 2, 0), firstMove.Offset(lastR, 1)).Copy _
         destination:=firstMove.Offset(advanceR, 0)
      firstMove.Range("A1:B2").Offset(lastR - 1, 0).ClearContents
      lastR = lastR - 2
      advanceR = advanceR + 1
End If
End If
advanceR = advanceR - 1
Loop

advanceR = lastR
Do While advanceR >= 0
If firstMove.Offset(advanceR + 2, 0).Value <> vbNullString Then
If firstMove.Offset(advanceR, 0).Value = firstMove.Offset(advanceR + 1, 0).Value And _
   firstMove.Offset(advanceR, 0).Value = firstMove.Offset(advanceR + 2, 0).Value And _
   firstMove.Offset(advanceR, 1).Value = firstMove.Offset(advanceR + 1, 1).Value And _
   firstMove.Offset(advanceR, 1).Value = firstMove.Offset(advanceR + 2, 1).Value Then
      If advanceR + 2 < lastR Then Range(firstMove.Offset(advanceR + 3, 0), firstMove.Offset(lastR, 1)).Copy _
         destination:=firstMove.Offset(advanceR + 1, 0)
      firstMove.Range("A1:B2").Offset(lastR - 1, 0).ClearContents
      lastR = lastR - 2
      firstMove.Offset(advanceR, 1).Value = -firstMove.Offset(advanceR, 1).Value
      advanceR = advanceR + 1
End If
End If
advanceR = advanceR - 1
Loop

Loop Until lastR = lastR0
End Sub

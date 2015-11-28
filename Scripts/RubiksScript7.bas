Attribute VB_Name = "Module7"
Sub simplify2(firstMove As Range)
Dim lastR As Integer, pointR As Integer, advanceR As Integer
Dim targetSide As Integer
Dim counter0 As Integer, counter1 As Integer
If firstMove.Value = vbNullString Then
    lastR = -1
ElseIf firstMove.Offset(1, 0).Value = vbNullString Then
    lastR = 0
Else
    lastR = firstMove.End(xlDown).Row - firstMove.Row
End If

For pointR = lastR To 0 Step -1
If firstMove.Offset(pointR, 0).Value Mod 3 <> 0 Then
    targetSide = 3 - firstMove.Offset(pointR, 0).Value Mod 3
    For advanceR = pointR - 1 To 0 Step -1
        If firstMove.Offset(advanceR, 0).Value Mod 3 = 0 Then
            Call sendToBack(advanceR, pointR, firstMove)
            Exit For
        ElseIf firstMove.Offset(advanceR, 0).Value Mod 3 = targetSide Then
            Call sendToBack(advanceR, pointR, firstMove)
            Call sendToBack(pointR - 1, pointR, firstMove)
            Exit For
        End If
    Next advanceR
    If advanceR = -1 Then Exit For
End If
Next pointR

targetSide = 3 - targetSide
If pointR > -1 Then counter1 = (pointR + 1) * (1 + 2 / 3 * targetSide) - 2 / 3 * WorksheetFunction.Sum(Range(firstMove, firstMove.Offset(pointR, 0)))
If lastR > pointR Then counter0 = (lastR - pointR) - 2 / 3 * WorksheetFunction.Sum(Range(firstMove.Offset(pointR + 1, 0), firstMove.Offset(lastR, 0)))
If lastR > -1 Then Range(firstMove, firstMove.Offset(lastR, 1)).ClearContents
Select Case counter1 Mod 4
Case 1, -3
    firstMove.Value = targetSide
    firstMove.Offset(0, 1).Value = 0
    pointR = 1
Case 2, -2
    firstMove.Range("A1:A2").Value = targetSide
    firstMove.Range("A1:A2").Offset(0, 1).Value = 0
    pointR = 2
Case 3, -1
    firstMove.Value = targetSide + 3
    firstMove.Offset(0, 1).Value = 0
    pointR = 1
Case 0
    pointR = 0
End Select
Select Case counter0 Mod 4
Case 1, -3
    firstMove.Offset(pointR, 0).Value = 0
    firstMove.Offset(pointR, 1).Value = 0
Case 2, -2
    If pointR < 2 Then firstMove.Range("A1:A2").Offset(pointR, 0).Value = 0 _
    Else firstMove.Range("A1:A2").Value = 3 - targetSide
    firstMove.Range("A1:A2").Offset(pointR, 1).Value = 0
Case 3, -1
    firstMove.Offset(pointR, 0).Value = 3
    firstMove.Offset(pointR, 1).Value = 0
End Select

End Sub


Sub sendToBack(fromR As Integer, toR As Integer, firstMove As Range)
Dim targetMove As Integer
targetMove = firstMove.Offset(fromR, 0).Value

Dim n As Integer
For i = fromR To toR - 1
n = (firstMove.Offset(i + 1, 0).Value + 6 - targetMove) Mod 3

If targetMove Mod 2 = 0 Then n = -n
Select Case n
   Case 0
      firstMove.Offset(i, 0).Value = firstMove.Offset(i + 1, 0).Value
   Case 1
      firstMove.Offset(i, 0).Value = (firstMove.Offset(i + 1, 0).Value + 1) Mod 6
   Case 2
      firstMove.Offset(i, 0).Value = (firstMove.Offset(i + 1, 0).Value + 2) Mod 6
   Case -1
      firstMove.Offset(i, 0).Value = (firstMove.Offset(i + 1, 0).Value + 4) Mod 6
   Case -2
      firstMove.Offset(i, 0).Value = (firstMove.Offset(i + 1, 0).Value + 5) Mod 6
End Select
firstMove.Offset(i, 1).Value = firstMove.Offset(i + 1, 1).Value
Next i

firstMove.Offset(toR, 0).Value = targetMove
firstMove.Offset(toR, 1).Value = 0
End Sub


Sub trial()
Dim testSize As Integer
testSize = 101

For i = 1 To testSize
Sheets("Simplify").Range("C" & i).Value = Int(6 * Rnd)
Next i
Sheets("Simplify").Range("D1:D" & testSize).Value = 0
Sheets("Simplify").Range("A1:B" & testSize).Value = Sheets("Simplify").Range("C1:D" & testSize).Value
Call test11(testSize)
Call simplify2(Sheets("Simplify").Range("A1"))
End Sub

Sub test11(testSize As Integer)
Call resetCUBE
For k = 1 To testSize
Call rotate(Sheets("Simplify").Range("C" & k).Value, True, 3, Sheets("Simplify").Range("C" & k))
Next k
Call rePaint(True)
End Sub

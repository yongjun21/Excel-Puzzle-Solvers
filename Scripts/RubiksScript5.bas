Attribute VB_Name = "Module5"
Sub randomize()
Sheets("Moves").Range("A3:N1002").Copy destination:=Sheets("Moves").Range("D3")
Sheets("Moves").Range("A3:B1002").Clear
Call resetCUBE

Dim nextMove As Range
Set nextMove = Sheets("Moves").Range("A3")
For k = 0 To 29
Call rotate(Int(6 * Rnd), Rnd > 0.5, 1, nextMove)
Next k
For k = 0 To 4
Call rotate(Int(6 * Rnd), Rnd > 0.5, 3, nextMove)
Next k
Call rePaint(True)
End Sub

Sub executeMoves(moveStart As Range)
If moveStart.Value = vbNullString Then Exit Sub Else Sheets("Main").Select
Dim targetMove As Range
Set targetMove = moveStart
Dim clockwise As Boolean, layer As Integer

Do While targetMove.Value <> vbNullString
   Application.Wait Now() + TimeValue("0:00:01")
   Select Case targetMove.Offset(0, 1).Value
   Case -1
      clockwise = False
      layer = 1
   Case 0
      clockwise = True
      layer = 3
   Case 1
      clockwise = True
      layer = 1
   End Select
   Call rotate(targetMove.Value, clockwise, layer)
   Set targetMove = targetMove.Offset(1, 0)
Loop
End Sub

Sub reverse(moveStart As Range)
If moveStart.Value = vbNullString Then Exit Sub
Sheets("Moves").Range("A3:N1002").Copy destination:=Sheets("Moves").Range("D3")
Dim k As Integer, kmax As Integer
If moveStart.Offset(1, 0).Value = vbNullString Then kmax = 0 Else kmax = moveStart.End(xlDown).Row - 3

For k = 0 To kmax
If moveStart.Offset(kmax - k, 4).Value = 0 Then
   Sheets("Moves").Range("A3").Offset(k, 0).Value = (moveStart.Offset(kmax - k, 3).Value + 3) Mod 6
   Sheets("Moves").Range("A3").Offset(k, 1).Value = 0
Else
   Sheets("Moves").Range("A3").Offset(k, 0).Value = moveStart.Offset(kmax - k, 3).Value
   Sheets("Moves").Range("A3").Offset(k, 1).Value = -moveStart.Offset(kmax - k, 4).Value
End If
Next k
End Sub

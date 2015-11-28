Attribute VB_Name = "Module3"
Sub rotate(side As Integer, clockwise As Boolean, layer As Integer, Optional nextMove As Range)
Dim i As Integer, j As Integer, ii As Integer, jj As Integer
Dim cw As Integer
For i = 0 To 5
For j = 0 To 8
ii = i
jj = j
Call transformAddress(ii, jj, side, clockwise, layer)
Sheets("Mapping").Range("F2").Offset(ii * 9 + jj, 0).Value = Sheets("Mapping").Range("F2").Offset(i * 9 + j, -1).Value
Next j
Next i
If nextMove Is Nothing Then
   Call rePaint(False)
Else
   Sheets("Mapping").Range("F2:F55").Offset(0, -1).Value = Sheets("Mapping").Range("F2:F55").Value
   If layer = 3 And Not clockwise Then side = (side + 3) Mod 6
   nextMove.Value = side
   If layer = 3 Then
      cw = 0
   ElseIf clockwise Then
      cw = 1
   Else
      cw = -1
   End If
   nextMove.Offset(0, 1).Value = cw
   Set nextMove = nextMove.Offset(1, 0)
End If
End Sub


Sub transformAddress(i As Integer, j As Integer, side As Integer, clockwise As Boolean, layer As Integer)
If layer = 1 And Sheets("Mapping").Range("A2").Offset(i * 9 + j, 1 + (side + 2) Mod 3).Value <> (side Mod 2) * 2 + 1 Then Exit Sub
If layer = 2 And Sheets("Mapping").Range("A2").Offset(i * 9 + j, 1 + (side + 2) Mod 3).Value <> 2 Then Exit Sub

Dim cw As Boolean
If side Mod 2 = 0 Then cw = clockwise Else cw = Not clockwise
Dim x As Integer, y As Integer, z As Integer, n As Integer
x = Sheets("Mapping").Range("A2").Offset(i * 9 + j, 1).Value
y = Sheets("Mapping").Range("A2").Offset(i * 9 + j, 2).Value
z = Sheets("Mapping").Range("A2").Offset(i * 9 + j, 3).Value

n = (i + 6 - side) Mod 3
If Not cw Then n = -n
Select Case n
   Case 1
      i = (i + 1) Mod 6
   Case 2
      i = (i + 2) Mod 6
   Case -1
      i = (i + 4) Mod 6
   Case -2
      i = (i + 5) Mod 6
End Select

Select Case side
   Case 0, 3
      Call rotateAddress(x, y, cw)
   Case 1, 4
      Call rotateAddress(y, z, cw)
   Case 2, 5
      Call rotateAddress(z, x, cw)
End Select

Select Case i
   Case 0, 3
      j = (y - 1) * 3 + x - 1
   Case 1, 4
      j = (z - 1) * 3 + y - 1
   Case 2, 5
      j = (x - 1) * 3 + z - 1
End Select
End Sub

Sub rotateAddress(a As Integer, b As Integer, clockwise As Boolean)
a = a - 2
b = b - 2
If clockwise Then
   If a = b Then
      b = -b
   ElseIf a = 0 Then
      a = b
      b = 0
   ElseIf b = 0 Then
      b = -a
      a = 0
   Else
      a = -a
   End If
Else
   If a = b Then
      a = -a
   ElseIf a = 0 Then
      a = -b
      b = 0
   ElseIf b = 0 Then
      b = a
      a = 0
   Else
      b = -b
   End If
End If
a = a + 2
b = b + 2
End Sub

Sub test()
Call rotate(0, False, 2)
End Sub

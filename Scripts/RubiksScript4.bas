Attribute VB_Name = "Module4"
Sub executeRotate(i As Integer, j As Integer, layer As Integer)
Dim side As Integer, clockwise As Boolean
Dim ii As Integer, jj As Integer
ii = Range("B1").Value
jj = Range("B2").Value

If Now() - Range("B3").Value > 1 / 24 / 60 / 60 * 3 Then
   Range("B1").Value = i
   Range("B2").Value = j
   Range("B3").Value = Now()
ElseIf i = ii And j = jj Then
   Range("B3").Value = 0
Else
   For side = 0 To 5
      Call transformAddress(ii, jj, side, True, layer)
      clockwise = True
      If ii = i And jj = j Then Exit For
      Call transformAddress(ii, jj, side, False, layer)
      Call transformAddress(ii, jj, side, False, layer)
      clockwise = False
      If ii = i And jj = j Then Exit For
      Call transformAddress(ii, jj, side, True, layer)
   Next side
   If side < 6 Then Call rotate(side, clockwise, layer)
   Range("B3").Value = 0
End If
End Sub

Sub button_00_click()
Call executeRotate(0, 0, 1)
End Sub

Sub button_01_click()
Call executeRotate(0, 1, 2)
End Sub

Sub button_02_click()
Call executeRotate(0, 2, 1)
End Sub

Sub button_03_click()
Call executeRotate(0, 3, 2)
End Sub

Sub button_04_click()
Call executeRotate(0, 4, 3)
End Sub

Sub button_05_click()
Call executeRotate(0, 5, 2)
End Sub

Sub button_06_click()
Call executeRotate(0, 6, 1)
End Sub

Sub button_07_click()
Call executeRotate(0, 7, 2)
End Sub

Sub button_08_click()
Call executeRotate(0, 8, 1)
End Sub

Sub button_20_click()
Call executeRotate(2, 0, 1)
End Sub

Sub button_21_click()
Call executeRotate(2, 1, 2)
End Sub

Sub button_22_click()
Call executeRotate(2, 2, 1)
End Sub

Sub button_23_click()
Call executeRotate(2, 3, 2)
End Sub

Sub button_24_click()
Call executeRotate(2, 4, 3)
End Sub

Sub button_25_click()
Call executeRotate(2, 5, 2)
End Sub

Sub button_26_click()
Call executeRotate(2, 6, 1)
End Sub

Sub button_27_click()
Call executeRotate(2, 7, 2)
End Sub

Sub button_28_click()
Call executeRotate(2, 8, 1)
End Sub

Sub button_40_click()
Call executeRotate(4, 0, 1)
End Sub

Sub button_41_click()
Call executeRotate(4, 1, 2)
End Sub

Sub button_42_click()
Call executeRotate(4, 2, 1)
End Sub

Sub button_43_click()
Call executeRotate(4, 3, 2)
End Sub

Sub button_44_click()
Call executeRotate(4, 4, 3)
End Sub

Sub button_45_click()
Call executeRotate(4, 5, 2)
End Sub

Sub button_46_click()
Call executeRotate(4, 6, 1)
End Sub

Sub button_47_click()
Call executeRotate(4, 7, 2)
End Sub

Sub button_48_click()
Call executeRotate(4, 8, 1)
End Sub

Sub button_51_click()
Call executeMoves(Sheets("Moves").Range("A3"))
End Sub

Sub button_52_click()
Call executeMoves(Sheets("Moves").Range("D3"))
End Sub

Sub button_53_click()
Call executeMoves(Sheets("Moves").Range("G3"))
End Sub

Sub button_54_click()
Call executeMoves(Sheets("Moves").Range("J3"))
End Sub

Sub button_55_click()
Call executeMoves(Sheets("Moves").Range("M3"))
End Sub

Sub button_61_click()
Call reverse(Sheets("Moves").Range("A3"))
End Sub

Sub button_62_click()
Call reverse(Sheets("Moves").Range("D3"))
End Sub

Sub button_63_click()
Call reverse(Sheets("Moves").Range("G3"))
End Sub

Sub button_64_click()
Call reverse(Sheets("Moves").Range("J3"))
End Sub

Sub button_65_click()
Call reverse(Sheets("Moves").Range("M3"))
End Sub

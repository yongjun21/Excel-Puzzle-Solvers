Attribute VB_Name = "Module2"
Sub rePaint(repaintALL As Boolean)
Dim location As Variant, colour As Long
Application.ScreenUpdating = False
For i = 1 To 54
If repaintALL Or Sheets("Mapping").Range("F1").Offset(i, 0).Value <> Sheets("Mapping").Range("F1").Offset(i, -1).Value Then
   colour = RGBshade(Sheets("Mapping").Range("F1").Offset(i, 0).Value)
   For j = 1 To 4
   If Sheets("Mapping").Range("F1").Offset(i, j).Value <> vbNullString Then
     location = Split(Sheets("Mapping").Range("F1").Offset(i, j).Value, ",")
     Call paintFace(Sheets("Main").Range("TA381").Offset(location(0), location(1)), Int(location(2)), colour)
   End If
   Next j
   Sheets("Mapping").Range("F1").Offset(i, -1).Value = Sheets("Mapping").Range("F1").Offset(i, 0).Value
End If
Next i
Application.ScreenUpdating = True
End Sub

Sub paintFace(origin As Range, direction As Integer, colour As Long)

Dim coordinates(18) As Variant
coordinates(0) = Array(1, 2, 2, 36)
coordinates(1) = Array(3, 3, 4, 35)
coordinates(2) = Array(5, 4, 6, 34)
coordinates(3) = Array(7, 5, 8, 33)
coordinates(4) = Array(9, 6, 10, 32)
coordinates(5) = Array(11, 7, 13, 31)
coordinates(6) = Array(14, 8, 15, 30)
coordinates(7) = Array(16, 9, 17, 29)
coordinates(8) = Array(18, 10, 19, 28)
coordinates(9) = Array(20, 11, 21, 27)
coordinates(10) = Array(22, 12, 23, 26)
coordinates(11) = Array(24, 13, 25, 25)
coordinates(12) = Array(26, 14, 28, 24)
coordinates(13) = Array(29, 15, 30, 23)
coordinates(14) = Array(31, 16, 32, 22)
coordinates(15) = Array(33, 17, 34, 21)
coordinates(16) = Array(35, 18, 36, 20)
coordinates(17) = Array(37, 19, 38, 19)

If direction = 1 Then
   For i = 0 To 17
      Range(origin.Offset(coordinates(i)(1) - 19, coordinates(i)(0) - 40), _
         origin.Offset(coordinates(i)(3) - 19, coordinates(i)(2) - 40)).Interior.color = colour
      Range(origin.Offset(coordinates(i)(1) - 38, -coordinates(i)(0)), _
         origin.Offset(coordinates(i)(3) - 38, -coordinates(i)(2))).Interior.color = colour
   Next i
   Range(origin.Offset(-18, -39), origin.Offset(-1, -1)).Interior.color = colour
ElseIf direction = 2 Then
   For i = 0 To 17
      Range(origin.Offset(coordinates(i)(1) - 19, -coordinates(i)(0) + 40), _
         origin.Offset(coordinates(i)(3) - 19, -coordinates(i)(2) + 40)).Interior.color = colour
      Range(origin.Offset(coordinates(i)(1) - 38, coordinates(i)(0)), _
         origin.Offset(coordinates(i)(3) - 38, coordinates(i)(2))).Interior.color = colour
   Next i
   Range(origin.Offset(-18, 1), origin.Offset(-1, 39)).Interior.color = colour
Else
   For i = 0 To 17
      Range(origin.Offset(coordinates(i)(1), coordinates(i)(0)), _
         origin.Offset(coordinates(i)(3), coordinates(i)(2))).Interior.color = colour
      Range(origin.Offset(coordinates(i)(1), -coordinates(i)(0)), _
         origin.Offset(coordinates(i)(3), -coordinates(i)(2))).Interior.color = colour
   Next i
   Range(origin.Offset(1, 0), origin.Offset(37, 0)).Interior.color = colour
End If
End Sub

Function RGBshade(colorCode As Integer) As Long
Select Case colorCode
   Case 0
      RGBshade = RGB(255, 255, 0)
   Case 1
      RGBshade = RGB(255, 127, 0)
   Case 2
      RGBshade = RGB(0, 0, 255)
   Case 3
      RGBshade = RGB(255, 255, 255)
   Case 4
      RGBshade = RGB(255, 0, 0)
   Case 5
      RGBshade = RGB(0, 255, 0)
End Select
End Function

Sub resetCUBE()
Sheets("Mapping").Range("F2:F55").Value = Sheets("Mapping").Range("A2:A55").Value
Call rePaint(True)
End Sub

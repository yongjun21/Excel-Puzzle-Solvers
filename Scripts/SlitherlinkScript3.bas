Attribute VB_Name = "Module3"
Function JordanCurve() As Boolean

Dim first As Integer, last As Integer
first = remainingEdge()

Do
last = unshaded()

For i = 2 To 14 Step 2
For j = 2 To 14 Step 2
   
   If shading(i, j) = 0 Then
      If shading(i - 2, j) <> 0 And topleft.Offset(i - 1, j) = 5 Then shading(i, j) = -shading(i - 2, j)
      If shading(i + 2, j) <> 0 And topleft.Offset(i + 1, j) = 5 Then shading(i, j) = -shading(i + 2, j)
      If shading(i, j - 2) <> 0 And topleft.Offset(i, j - 1) = 5 Then shading(i, j) = -shading(i, j - 2)
      If shading(i, j + 2) <> 0 And topleft.Offset(i, j + 1) = 5 Then shading(i, j) = -shading(i, j + 2)
      If shading(i - 2, j) <> 0 And topleft.Offset(i - 1, j) = -1 Then shading(i, j) = shading(i - 2, j)
      If shading(i + 2, j) <> 0 And topleft.Offset(i + 1, j) = -1 Then shading(i, j) = shading(i + 2, j)
      If shading(i, j - 2) <> 0 And topleft.Offset(i, j - 1) = -1 Then shading(i, j) = shading(i, j - 2)
      If shading(i, j + 2) <> 0 And topleft.Offset(i, j + 1) = -1 Then shading(i, j) = shading(i, j + 2)
   Else
      If topleft.Offset(i - 1, j).Value = 0 And shading(i, j) = -shading(i - 2, j) Then _
         Call confirm(topleft.Offset(i - 1, j), True)
      If topleft.Offset(i + 1, j).Value = 0 And shading(i, j) = -shading(i + 2, j) Then _
         Call confirm(topleft.Offset(i + 1, j), True)
      If topleft.Offset(i, j - 1).Value = 0 And shading(i, j) = -shading(i, j - 2) Then _
         Call confirm(topleft.Offset(i, j - 1), False)
      If topleft.Offset(i, j + 1).Value = 0 And shading(i, j) = -shading(i, j + 2) Then _
         Call confirm(topleft.Offset(i, j + 1), False)
      If topleft.Offset(i - 1, j).Value = 0 And shading(i, j) = shading(i - 2, j) Then _
         topleft.Offset(i - 1, j).Value = -1
      If topleft.Offset(i + 1, j).Value = 0 And shading(i, j) = shading(i + 2, j) Then _
         topleft.Offset(i + 1, j).Value = -1
      If topleft.Offset(i, j - 1).Value = 0 And shading(i, j) = shading(i, j - 2) Then _
         topleft.Offset(i, j - 1).Value = -1
      If topleft.Offset(i, j + 1).Value = 0 And shading(i, j) = shading(i, j + 2) Then _
         topleft.Offset(i, j + 1).Value = -1
   End If

   If origState(i, j) = 0 Then
      If shading(i - 2, j) <> 0 Then shading(i, j) = shading(i - 2, j)
      If shading(i + 2, j) <> 0 Then shading(i, j) = shading(i + 2, j)
      If shading(i, j - 2) <> 0 Then shading(i, j) = shading(i, j - 2)
      If shading(i, j + 2) <> 0 Then shading(i, j) = shading(i, j + 2)
      Call confirmShade(i, j, shading(i, j))
   End If
   
   If origState(i, j) = 1 Then
      If shading(i, j) <> 0 Then
         If shading(i - 2, j) = -shading(i, j) Then Call confirmShade(i, j, shading(i, j))
         If shading(i + 2, j) = -shading(i, j) Then Call confirmShade(i, j, shading(i, j))
         If shading(i, j - 2) = -shading(i, j) Then Call confirmShade(i, j, shading(i, j))
         If shading(i, j + 2) = -shading(i, j) Then Call confirmShade(i, j, shading(i, j))
      ElseIf nShade(i, j, -1) > 1 Then
         shading(i, j) = -1
      ElseIf nShade(i, j, 1) > 1 Then
         shading(i, j) = 1
      End If
   End If
   
   If origState(i, j) = 2 Then
      If nShade(i, j, -1) = 2 Then Call confirmShade(i, j, 1)
      If nShade(i, j, 1) = 2 Then Call confirmShade(i, j, -1)
   End If
   
   If origState(i, j) = 3 Then
      If shading(i, j) <> 0 Then
         If shading(i - 2, j) = shading(i, j) Then Call confirmShade(i, j, -shading(i, j))
         If shading(i + 2, j) = shading(i, j) Then Call confirmShade(i, j, -shading(i, j))
         If shading(i, j - 2) = shading(i, j) Then Call confirmShade(i, j, -shading(i, j))
         If shading(i, j + 2) = shading(i, j) Then Call confirmShade(i, j, -shading(i, j))
      ElseIf nShade(i, j, -1) > 1 Then
         shading(i, j) = 1
      ElseIf nShade(i, j, 1) > 1 Then
         shading(i, j) = -1
      End If
   End If

Next j
Next i

Range("BD38:BT54").Value = shading

zeros = 0
For i = 2 To 14 Step 2
For j = 2 To 14 Step 2
   If shading(i, j) = 0 Then zeros = zeros + 1
Next j
Next i

Loop Until unshaded() = last

JordanCurve = (remainingEdge() <> first)

End Function


Function nShade(i, j, shade As Integer) As Integer
If shading(i - 2, j) = shade Then nShade = nShade + 1
If shading(i + 2, j) = shade Then nShade = nShade + 1
If shading(i, j - 2) = shade Then nShade = nShade + 1
If shading(i, j + 2) = shade Then nShade = nShade + 1
End Function

Sub confirmShade(i, j, shade As Integer)
If shading(i - 2, j) = 0 Then shading(i - 2, j) = shade
If shading(i + 2, j) = 0 Then shading(i + 2, j) = shade
If shading(i, j - 2) = 0 Then shading(i, j - 2) = shade
If shading(i, j + 2) = 0 Then shading(i, j + 2) = shade
End Sub

Function unshaded() As Integer
For i = 2 To 14 Step 2
For j = 2 To 14 Step 2
   If shading(i, j) = 0 Then unshaded = unshaded + 1
Next j
Next i
End Function


Function randomSearch() As Boolean
Dim saveState As Variant
Dim saveCorner(0 To 16, 0 To 16, 1 To 4, 0 To 2) As Boolean
Dim HorV As Boolean
Dim i As Integer, j As Integer
Dim result As Boolean

For a = 1 To 15 Step 2
For b = 2 To 14 Step 2
   HorV = True
   i = a
   j = b
   If topleft.Offset(i, j).Value = 0 Then Exit For
   HorV = False
   i = b
   j = a
   If topleft.Offset(i, j).Value = 0 Then Exit For
Next b
If b <= 14 Then Exit For
Next a
If a > 15 Then Exit Function

saveState = topleft.Range("A1:Q17").Value
Call copyArray(corner, saveCorner)
topleft.Offset(i, j).Value = -1
Do
   change = narrowPlus()
Loop Until Not change Or Not unfinished()
If Not unfinished() Then
   randomSearch = True
   Exit Function
ElseIf randomSearch() Then
   randomSearch = True
   Exit Function
End If
   
topleft.Range("A1:Q17").Value = saveState
Call copyArray(saveCorner, corner)
Call confirm(topleft.Offset(i, j), HorV)
Do
   change = narrowPlus()
Loop Until Not change Or Not unfinished()
If Not unfinished() Then
   randomSearch = True
   Exit Function
ElseIf randomSearch() Then
   randomSearch = True
   Exit Function
End If

End Function


Sub copyArray(sourceArray() As Boolean, destArray() As Boolean)
For i = 1 To 15 Step 2
For j = 1 To 15 Step 2
For d = 1 To 4
For n = 0 To 2
   destArray(i, j, d, n) = sourceArray(i, j, d, n)
Next n
Next d
Next j
Next i
End Sub



Attribute VB_Name = "Module2"
Function narrowPlus() As Boolean

Dim first As Integer, last As Integer
first = remainingEdge()

Do
last = eliminated()

For i = 1 To 15 Step 2
For j = 1 To 15 Step 2
   'if at least one x then not 2 (if 2 then no x)
   corner(i, j, 1, 2) = (corner(i, j, 1, 2) Or topleft.Offset(i - 1, j).Value = -1 Or topleft.Offset(i, j - 1).Value = -1)
   corner(i, j, 2, 2) = (corner(i, j, 2, 2) Or topleft.Offset(i - 1, j).Value = -1 Or topleft.Offset(i, j + 1).Value = -1)
   corner(i, j, 3, 2) = (corner(i, j, 3, 2) Or topleft.Offset(i + 1, j).Value = -1 Or topleft.Offset(i, j - 1).Value = -1)
   corner(i, j, 4, 2) = (corner(i, j, 4, 2) Or topleft.Offset(i + 1, j).Value = -1 Or topleft.Offset(i, j + 1).Value = -1)
   
   'if at least one cf then not 0 (if 0 then no cf)
   corner(i, j, 1, 0) = (corner(i, j, 1, 0) Or topleft.Offset(i - 1, j).Value = 5 Or topleft.Offset(i, j - 1).Value = 5)
   corner(i, j, 2, 0) = (corner(i, j, 2, 0) Or topleft.Offset(i - 1, j).Value = 5 Or topleft.Offset(i, j + 1).Value = 5)
   corner(i, j, 3, 0) = (corner(i, j, 3, 0) Or topleft.Offset(i + 1, j).Value = 5 Or topleft.Offset(i, j - 1).Value = 5)
   corner(i, j, 4, 0) = (corner(i, j, 4, 0) Or topleft.Offset(i + 1, j).Value = 5 Or topleft.Offset(i, j + 1).Value = 5)
   
   'if both x or both cf then not 1
   corner(i, j, 1, 1) = (corner(i, j, 1, 1) Or (topleft.Offset(i - 1, j).Value = -1 And topleft.Offset(i, j - 1).Value = -1) Or (topleft.Offset(i - 1, j).Value = 5 And topleft.Offset(i, j - 1).Value = 5))
   corner(i, j, 2, 1) = (corner(i, j, 2, 1) Or (topleft.Offset(i - 1, j).Value = -1 And topleft.Offset(i, j + 1).Value = -1) Or (topleft.Offset(i - 1, j).Value = 5 And topleft.Offset(i, j + 1).Value = 5))
   corner(i, j, 3, 1) = (corner(i, j, 3, 1) Or (topleft.Offset(i + 1, j).Value = -1 And topleft.Offset(i, j - 1).Value = -1) Or (topleft.Offset(i + 1, j).Value = 5 And topleft.Offset(i, j - 1).Value = 5))
   corner(i, j, 4, 1) = (corner(i, j, 4, 1) Or (topleft.Offset(i + 1, j).Value = -1 And topleft.Offset(i, j + 1).Value = -1) Or (topleft.Offset(i + 1, j).Value = 5 And topleft.Offset(i, j + 1).Value = 5))
   
   'if opp not 0 then not 2 (if 2 then opp 0)
   corner(i, j, 4, 2) = (corner(i, j, 4, 2) Or corner(i, j, 1, 0))
   corner(i, j, 3, 2) = (corner(i, j, 3, 2) Or corner(i, j, 2, 0))
   corner(i, j, 2, 2) = (corner(i, j, 2, 2) Or corner(i, j, 3, 0))
   corner(i, j, 1, 2) = (corner(i, j, 1, 2) Or corner(i, j, 4, 0))
   
   'if opp not 1 then not 1 (if 1 then opp 1)
   corner(i, j, 4, 1) = (corner(i, j, 4, 1) Or corner(i, j, 1, 1))
   corner(i, j, 3, 1) = (corner(i, j, 3, 1) Or corner(i, j, 2, 1))
   corner(i, j, 2, 1) = (corner(i, j, 2, 1) Or corner(i, j, 3, 1))
   corner(i, j, 1, 1) = (corner(i, j, 1, 1) Or corner(i, j, 4, 1))
   
   'if opp not 0 and not 2 then not 0 (if 0 then opp not 1)
   corner(i, j, 4, 0) = (corner(i, j, 4, 0) Or (corner(i, j, 1, 0) And corner(i, j, 1, 2)))
   corner(i, j, 3, 0) = (corner(i, j, 3, 0) Or (corner(i, j, 2, 0) And corner(i, j, 2, 2)))
   corner(i, j, 2, 0) = (corner(i, j, 2, 0) Or (corner(i, j, 3, 0) And corner(i, j, 3, 2)))
   corner(i, j, 1, 0) = (corner(i, j, 1, 0) Or (corner(i, j, 4, 0) And corner(i, j, 4, 2)))
Next j
Next i

For i = 2 To 14 Step 2
For j = 2 To 14 Step 2
   If origState(i, j) = 1 Then
      'if opp not 1 then not 0 (if 0 then opp 1)
      corner(i - 1, j - 1, 4, 0) = (corner(i - 1, j - 1, 4, 0) Or corner(i + 1, j + 1, 1, 1))
      corner(i - 1, j + 1, 3, 0) = (corner(i - 1, j + 1, 3, 0) Or corner(i + 1, j - 1, 2, 1))
      corner(i + 1, j - 1, 2, 0) = (corner(i + 1, j - 1, 2, 0) Or corner(i - 1, j + 1, 3, 1))
      corner(i + 1, j + 1, 1, 0) = (corner(i + 1, j + 1, 1, 0) Or corner(i - 1, j - 1, 4, 1))
      
      'if opp not 0 then not 1 (if 1 then opp 0)
      corner(i - 1, j - 1, 4, 1) = (corner(i - 1, j - 1, 4, 1) Or corner(i + 1, j + 1, 1, 0))
      corner(i - 1, j + 1, 3, 1) = (corner(i - 1, j + 1, 3, 1) Or corner(i + 1, j - 1, 2, 0))
      corner(i + 1, j - 1, 2, 1) = (corner(i + 1, j - 1, 2, 1) Or corner(i - 1, j + 1, 3, 0))
      corner(i + 1, j + 1, 1, 1) = (corner(i + 1, j + 1, 1, 1) Or corner(i - 1, j - 1, 4, 0))

   ElseIf origState(i, j) = 3 Then
      'if opp not 1 then not 2 (if 2 then opp 1)
      corner(i - 1, j - 1, 4, 2) = (corner(i - 1, j - 1, 4, 2) Or corner(i + 1, j + 1, 1, 1))
      corner(i - 1, j + 1, 3, 2) = (corner(i - 1, j + 1, 3, 2) Or corner(i + 1, j - 1, 2, 1))
      corner(i + 1, j - 1, 2, 2) = (corner(i + 1, j - 1, 2, 2) Or corner(i - 1, j + 1, 3, 1))
      corner(i + 1, j + 1, 1, 2) = (corner(i + 1, j + 1, 1, 2) Or corner(i - 1, j - 1, 4, 1))
      
      'if opp not 2 then not 1 (if 1 then opp 2)
      corner(i - 1, j - 1, 4, 1) = (corner(i - 1, j - 1, 4, 1) Or corner(i + 1, j + 1, 1, 2))
      corner(i - 1, j + 1, 3, 1) = (corner(i - 1, j + 1, 3, 1) Or corner(i + 1, j - 1, 2, 2))
      corner(i + 1, j - 1, 2, 1) = (corner(i + 1, j - 1, 2, 1) Or corner(i - 1, j + 1, 3, 2))
      corner(i + 1, j + 1, 1, 1) = (corner(i + 1, j + 1, 1, 1) Or corner(i - 1, j - 1, 4, 2))

   ElseIf origState(i, j) = 2 Then
      'if opp not 2 then not 0 (if 0 then opp 2)
      corner(i - 1, j - 1, 4, 0) = (corner(i - 1, j - 1, 4, 0) Or corner(i + 1, j + 1, 1, 2))
      corner(i - 1, j + 1, 3, 0) = (corner(i - 1, j + 1, 3, 0) Or corner(i + 1, j - 1, 2, 2))
      corner(i + 1, j - 1, 2, 0) = (corner(i + 1, j - 1, 2, 0) Or corner(i - 1, j + 1, 3, 2))
      corner(i + 1, j + 1, 1, 0) = (corner(i + 1, j + 1, 1, 0) Or corner(i - 1, j - 1, 4, 2))
      
      'if vopp not 1 then not 0 (if 0 then vopp 1)
      corner(i - 1, j - 1, 4, 0) = (corner(i - 1, j - 1, 4, 0) Or corner(i - 1, j + 1, 3, 1))
      corner(i - 1, j + 1, 3, 0) = (corner(i - 1, j + 1, 3, 0) Or corner(i - 1, j - 1, 4, 1))
      corner(i + 1, j - 1, 2, 0) = (corner(i + 1, j - 1, 2, 0) Or corner(i + 1, j + 1, 1, 1))
      corner(i + 1, j + 1, 1, 0) = (corner(i + 1, j + 1, 1, 0) Or corner(i + 1, j - 1, 2, 1))
      
      'if hopp not 1 then not 0 (if 0 then hopp 1)
      corner(i - 1, j - 1, 4, 0) = (corner(i - 1, j - 1, 4, 0) Or corner(i + 1, j - 1, 2, 1))
      corner(i - 1, j + 1, 3, 0) = (corner(i - 1, j + 1, 3, 0) Or corner(i + 1, j + 1, 1, 1))
      corner(i + 1, j - 1, 2, 0) = (corner(i + 1, j - 1, 2, 0) Or corner(i - 1, j - 1, 4, 1))
      corner(i + 1, j + 1, 1, 0) = (corner(i + 1, j + 1, 1, 0) Or corner(i - 1, j + 1, 3, 1))
      
      'if opp not 0 then not 2 (if 2 then opp 0)
      corner(i - 1, j - 1, 4, 2) = (corner(i - 1, j - 1, 4, 2) Or corner(i + 1, j + 1, 1, 0))
      corner(i - 1, j + 1, 3, 2) = (corner(i - 1, j + 1, 3, 2) Or corner(i + 1, j - 1, 2, 0))
      corner(i + 1, j - 1, 2, 2) = (corner(i + 1, j - 1, 2, 2) Or corner(i - 1, j + 1, 3, 0))
      corner(i + 1, j + 1, 1, 2) = (corner(i + 1, j + 1, 1, 2) Or corner(i - 1, j - 1, 4, 0))
      
      'if vopp not 1 then not 2 (if 2 then vopp 1)
      corner(i - 1, j - 1, 4, 2) = (corner(i - 1, j - 1, 4, 2) Or corner(i - 1, j + 1, 3, 1))
      corner(i - 1, j + 1, 3, 2) = (corner(i - 1, j + 1, 3, 2) Or corner(i - 1, j - 1, 4, 1))
      corner(i + 1, j - 1, 2, 2) = (corner(i + 1, j - 1, 2, 2) Or corner(i + 1, j + 1, 1, 1))
      corner(i + 1, j + 1, 1, 2) = (corner(i + 1, j + 1, 1, 2) Or corner(i + 1, j - 1, 2, 1))
      
      'if hopp not 1 then not 2 (if 2 then hopp 1)
      corner(i - 1, j - 1, 4, 2) = (corner(i - 1, j - 1, 4, 2) Or corner(i + 1, j - 1, 2, 1))
      corner(i - 1, j + 1, 3, 2) = (corner(i - 1, j + 1, 3, 2) Or corner(i + 1, j + 1, 1, 1))
      corner(i + 1, j - 1, 2, 2) = (corner(i + 1, j - 1, 2, 2) Or corner(i - 1, j - 1, 4, 1))
      corner(i + 1, j + 1, 1, 2) = (corner(i + 1, j + 1, 1, 2) Or corner(i - 1, j + 1, 3, 1))
      
      'if opp not 1 then not 1 (if 1 then opp 1)
      corner(i - 1, j - 1, 4, 1) = (corner(i - 1, j - 1, 4, 1) Or corner(i + 1, j + 1, 1, 1))
      corner(i - 1, j + 1, 3, 1) = (corner(i - 1, j + 1, 3, 1) Or corner(i + 1, j - 1, 2, 1))
      corner(i + 1, j - 1, 2, 1) = (corner(i + 1, j - 1, 2, 1) Or corner(i - 1, j + 1, 3, 1))
      corner(i + 1, j + 1, 1, 1) = (corner(i + 1, j + 1, 1, 1) Or corner(i - 1, j - 1, 4, 1))
      
   End If
Next j
Next i

Loop Until eliminated() = last

'Call printCorner()
Do
Do
last = remainingEdge()

For i = 1 To 15 Step 2
For j = 1 To 15 Step 2
If topleft.Offset(i, j).Value <> 5 Then
   If corner(i, j, 1, 0) And corner(i, j, 1, 1) Then 'is 2
      If topleft.Offset(i - 1, j).Value = 0 Then Call confirm(topleft.Offset(i - 1, j), False)
      If topleft.Offset(i, j - 1).Value = 0 Then Call confirm(topleft.Offset(i, j - 1), True)
   ElseIf corner(i, j, 1, 1) And corner(i, j, 1, 2) Then ' is 0
      topleft.Offset(i - 1, j).Value = -1
      topleft.Offset(i, j - 1).Value = -1
   ElseIf corner(i, j, 1, 0) And corner(i, j, 1, 2) Then ' is 1
      If topleft.Offset(i - 1, j).Value = 5 Then topleft.Offset(i, j - 1).Value = -1
      If topleft.Offset(i, j - 1).Value = 5 Then topleft.Offset(i - 1, j).Value = -1
      If topleft.Offset(i - 1, j).Value = 0 And topleft.Offset(i, j - 1).Value = -1 Then Call confirm(topleft.Offset(i - 1, j), False)
      If topleft.Offset(i, j - 1).Value = 0 And topleft.Offset(i - 1, j).Value = -1 Then Call confirm(topleft.Offset(i, j - 1), True)
   End If
   
   If corner(i, j, 2, 0) And corner(i, j, 2, 1) Then
      If topleft.Offset(i - 1, j).Value = 0 Then Call confirm(topleft.Offset(i - 1, j), False)
      If topleft.Offset(i, j + 1).Value = 0 Then Call confirm(topleft.Offset(i, j + 1), True)
   ElseIf corner(i, j, 2, 1) And corner(i, j, 2, 2) Then
      topleft.Offset(i - 1, j).Value = -1
      topleft.Offset(i, j + 1).Value = -1
   ElseIf corner(i, j, 2, 0) And corner(i, j, 2, 2) Then
      If topleft.Offset(i - 1, j).Value = 5 Then topleft.Offset(i, j + 1).Value = -1
      If topleft.Offset(i, j + 1).Value = 5 Then topleft.Offset(i - 1, j).Value = -1
      If topleft.Offset(i - 1, j).Value = 0 And topleft.Offset(i, j + 1).Value = -1 Then Call confirm(topleft.Offset(i - 1, j), False)
      If topleft.Offset(i, j + 1).Value = 0 And topleft.Offset(i - 1, j).Value = -1 Then Call confirm(topleft.Offset(i, j + 1), True)
   End If
   
   If corner(i, j, 3, 0) And corner(i, j, 3, 1) Then
      If topleft.Offset(i + 1, j).Value = 0 Then Call confirm(topleft.Offset(i + 1, j), False)
      If topleft.Offset(i, j - 1).Value = 0 Then Call confirm(topleft.Offset(i, j - 1), True)
   ElseIf corner(i, j, 3, 1) And corner(i, j, 3, 2) Then
      topleft.Offset(i + 1, j).Value = -1
      topleft.Offset(i, j - 1).Value = -1
   ElseIf corner(i, j, 3, 0) And corner(i, j, 3, 2) Then
      If topleft.Offset(i + 1, j).Value = 5 Then topleft.Offset(i, j - 1).Value = -1
      If topleft.Offset(i, j - 1).Value = 5 Then topleft.Offset(i + 1, j).Value = -1
      If topleft.Offset(i + 1, j).Value = 0 And topleft.Offset(i, j - 1).Value = -1 Then Call confirm(topleft.Offset(i + 1, j), False)
      If topleft.Offset(i, j - 1).Value = 0 And topleft.Offset(i + 1, j).Value = -1 Then Call confirm(topleft.Offset(i, j - 1), True)
   End If
   
   If corner(i, j, 4, 0) And corner(i, j, 4, 1) Then
      If topleft.Offset(i + 1, j).Value = 0 Then Call confirm(topleft.Offset(i + 1, j), False)
      If topleft.Offset(i, j + 1).Value = 0 Then Call confirm(topleft.Offset(i, j + 1), True)
   ElseIf corner(i, j, 4, 1) And corner(i, j, 4, 2) Then
      topleft.Offset(i + 1, j).Value = -1
      topleft.Offset(i, j + 1).Value = -1
   ElseIf corner(i, j, 4, 0) And corner(i, j, 4, 2) Then
      If topleft.Offset(i + 1, j).Value = 5 Then topleft.Offset(i, j + 1).Value = -1
      If topleft.Offset(i, j + 1).Value = 5 Then topleft.Offset(i + 1, j).Value = -1
      If topleft.Offset(i + 1, j).Value = 0 And topleft.Offset(i, j + 1).Value = -1 Then Call confirm(topleft.Offset(i + 1, j), False)
      If topleft.Offset(i, j + 1).Value = 0 And topleft.Offset(i + 1, j).Value = -1 Then Call confirm(topleft.Offset(i, j + 1), True)
   End If
End If
Next j
Next i

For i = 2 To 14 Step 2
For j = 2 To 14 Step 2
   If topleft.Offset(i, j).Value > 0 And nState(topleft.Offset(i, j), 0) = topleft.Offset(i, j).Value Then
      If topleft.Offset(i - 1, j).Value = 0 Then Call confirm(topleft.Offset(i - 1, j), True)
      If topleft.Offset(i + 1, j).Value = 0 Then Call confirm(topleft.Offset(i + 1, j), True)
      If topleft.Offset(i, j - 1).Value = 0 Then Call confirm(topleft.Offset(i, j - 1), False)
      If topleft.Offset(i, j + 1).Value = 0 Then Call confirm(topleft.Offset(i, j + 1), False)
      topleft.Offset(i, j).Value = -1
   End If
Next j
Next i

Loop Until remainingEdge() = last

Dim LooseEnds As Range, OpenEnd As Range, OtherEnd As Range
Set LooseEnds = findLooseEnds()
If LooseEnds.Areas.Count > 2 Then
   For Each OpenEnd In LooseEnds.Areas
      Set OtherEnd = trace(OpenEnd, 1, 4, vbNullString, 0)
      If OpenEnd.Offset(-2, 0).Address = OtherEnd.Address And OpenEnd.Offset(-1, 0).Value = 0 Then
         OpenEnd.Offset(-1, 0).Value = -1
      ElseIf OpenEnd.Offset(2, 0).Address = OtherEnd.Address And OpenEnd.Offset(1, 0).Value = 0 Then
         OpenEnd.Offset(1, 0).Value = -1
      ElseIf OpenEnd.Offset(0, -2).Address = OtherEnd.Address And OpenEnd.Offset(0, -1).Value = 0 Then
         OpenEnd.Offset(0, -1).Value = -1
      ElseIf OpenEnd.Offset(0, 2).Address = OtherEnd.Address And OpenEnd.Offset(0, 1).Value = 0 Then
         OpenEnd.Offset(0, 1).Value = -1
      End If
   Next
End If

Loop Until remainingEdge() = last

narrowPlus = (remainingEdge() <> first)

End Function


Sub confirm(edge As Range, HorV As Boolean)

edge.Value = 5

If HorV Then
   edge.Offset(-1, 0).Value = edge.Offset(-1, 0).Value - 1
   If edge.Offset(-1, 0).Value = 0 Then Call closeOpen(edge.Offset(-1, 0), False)
   
   edge.Offset(1, 0).Value = edge.Offset(1, 0).Value - 1
   If edge.Offset(1, 0).Value = 0 Then Call closeOpen(edge.Offset(1, 0), False)
   
   If edge.Offset(0, -1).Value = 1 Then Call closeOpen(edge.Offset(0, -1), True) _
      Else edge.Offset(0, -1).Value = 1
   
   If edge.Offset(0, 1).Value = 1 Then Call closeOpen(edge.Offset(0, 1), True) _
      Else edge.Offset(0, 1).Value = 1
      
Else
   If edge.Offset(-1, 0).Value = 1 Then Call closeOpen(edge.Offset(-1, 0), True) _
      Else edge.Offset(-1, 0).Value = 1
   
   If edge.Offset(1, 0).Value = 1 Then Call closeOpen(edge.Offset(1, 0), True) _
      Else edge.Offset(1, 0).Value = 1
   
   edge.Offset(0, -1).Value = edge.Offset(0, -1).Value - 1
   If edge.Offset(0, -1).Value = 0 Then Call closeOpen(edge.Offset(0, -1), False)
   
   edge.Offset(0, 1).Value = edge.Offset(0, 1).Value - 1
   If edge.Offset(0, 1).Value = 0 Then Call closeOpen(edge.Offset(0, 1), False)
   
End If

End Sub


Sub closeOpen(vertex As Range, color As Boolean)
If vertex.Offset(-1, 0).Value = 0 Then vertex.Offset(-1, 0).Value = -1
If vertex.Offset(1, 0).Value = 0 Then vertex.Offset(1, 0).Value = -1
If vertex.Offset(0, -1).Value = 0 Then vertex.Offset(0, -1).Value = -1
If vertex.Offset(0, 1).Value = 0 Then vertex.Offset(0, 1).Value = -1
If color Then vertex.Value = 5 Else vertex.Value = -1
End Sub


Function eliminated() As Integer
For i = 1 To 15 Step 2
For j = 1 To 15 Step 2
For d = 1 To 4
For n = 0 To 2
   If corner(i, j, d, n) Then eliminated = eliminated + 1
Next n
Next d
Next j
Next i
End Function

Function remainingEdge() As Integer
For i = 1 To 15 Step 2
For j = 2 To 14 Step 2
   If topleft.Offset(i, j) = 0 Then remainingEdge = remainingEdge + 1
Next j
Next i
For i = 2 To 14 Step 2
For j = 1 To 15 Step 2
   If topleft.Offset(i, j) = 0 Then remainingEdge = remainingEdge + 1
Next j
Next i
End Function

Sub printCorner()
For i = 1 To 15 Step 2
For j = 1 To 15 Step 2
For k = 1 To 4
For n = 0 To 2
   Sheets("Sheet3").Range("D2").Offset((i \ 2) * 8 + j \ 2, (k - 1) * 4 + n).Value = corner(i, j, k, n)
Next n
Next k
Next j
Next i
End Sub


Function findLooseEnds() As Range
Dim nonEmpty As Boolean
Set findLooseEnds = topleft.Offset(1, 1)
For i = 1 To 15 Step 2
For j = 1 To 15 Step 2
If topleft.Offset(i, j) = 1 Then
   If nonEmpty Then
      Set findLooseEnds = Union(findLooseEnds, topleft.Offset(i, j))
   Else
      Set findLooseEnds = topleft.Offset(i, j)
      nonEmpty = True
   End If
End If
Next j
Next i
End Function


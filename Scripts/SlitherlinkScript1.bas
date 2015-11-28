Attribute VB_Name = "Module1"
Public topleft As Range
Public origState(0 To 16, 0 To 16) As Integer
Public corner(0 To 16, 0 To 16, 1 To 4, 0 To 2) As Boolean
Public shading(0 To 16, 0 To 16) As Integer

Sub main()
Erase origState, corner, shading
Set topleft = Range("BD20")
Call redoFormatting

For i = 2 To 14 Step 2
For j = 2 To 14 Step 2
   If topleft.Offset(i - 18, j - 54).Value <> vbNullString Then _
      origState(i, j) = topleft.Offset(i - 18, j - 54).Value Else origState(i, j) = -1
Next j
Next i

For i = 2 To 14 Step 2
For j = 2 To 14 Step 2
   If origState(i, j) = 0 Then
      corner(i - 1, j - 1, 4, 1) = True
      corner(i - 1, j - 1, 4, 2) = True
      corner(i - 1, j + 1, 3, 1) = True
      corner(i - 1, j + 1, 3, 2) = True
      corner(i + 1, j - 1, 2, 1) = True
      corner(i + 1, j - 1, 2, 2) = True
      corner(i + 1, j + 1, 1, 1) = True
      corner(i + 1, j + 1, 1, 2) = True
   ElseIf origState(i, j) = 1 Then
      corner(i - 1, j - 1, 4, 2) = True
      corner(i - 1, j + 1, 3, 2) = True
      corner(i + 1, j - 1, 2, 2) = True
      corner(i + 1, j + 1, 1, 2) = True
   ElseIf origState(i, j) = 3 Then
      corner(i - 1, j - 1, 4, 0) = True
      corner(i - 1, j + 1, 3, 0) = True
      corner(i + 1, j - 1, 2, 0) = True
      corner(i + 1, j + 1, 1, 0) = True
   End If
Next j
Next i

For k = 0 To 16
   origState(0, k) = -1
   origState(16, k) = -1
   origState(k, 0) = -1
   origState(k, 16) = -1
   shading(0, k) = -1
   shading(16, k) = -1
   shading(k, 0) = -1
   shading(k, 16) = -1
Next k

topleft.Range("A1:Q17").Value = origState

Dim change As Boolean, change2 As Boolean
Do
   change = narrowPlus()
Loop Until Not change Or Not unfinished()

If Not unfinished() Then Exit Sub
Do
   change = JordanCurve() Or narrowPlus()
Loop Until Not change Or Not unfinished()

If Not unfinished() Then Exit Sub
Dim inconsistent As Boolean
For i = 1 To 15 Step 2
For j = 1 To 15 Step 2
   If corner(i, j, 1, 0) And corner(i, j, 1, 1) And corner(i, j, 1, 2) Then inconsistent = True
   If corner(i, j, 2, 0) And corner(i, j, 2, 1) And corner(i, j, 2, 2) Then inconsistent = True
   If corner(i, j, 3, 0) And corner(i, j, 3, 1) And corner(i, j, 3, 2) Then inconsistent = True
   If corner(i, j, 4, 0) And corner(i, j, 4, 1) And corner(i, j, 4, 2) Then inconsistent = True
Next j
Next i
If Not inconsistent Then
   If randomSearch() Then Exit Sub
End If
Call clearlines
MsgBox "No feasible solution"

End Sub


Function unfinished() As Boolean

Dim counter As Integer

For i = 1 To 15 Step 2
For j = 1 To 15 Step 2
   If topleft.Offset(i, j).Value = 1 Then
      unfinished = True
      Exit Function
   ElseIf topleft.Offset(i, j).Value = 5 Then
      counter = counter + 1
   End If
Next j
Next i

For i = 2 To 14 Step 2
For j = 2 To 14 Step 2
   If origState(i, j) <> -1 And origState(i, j) <> nState(topleft.Offset(i, j), 5) Then
      unfinished = True
      Exit Function
   End If
Next j
Next i

For i = 1 To 15 Step 2
For j = 2 To 14 Step 2
   If topleft.Offset(i, j) = 5 Then
      Set dummy = trace(topleft.Offset(i, j + 1), 1, 3, topleft.Offset(i, j - 1).Address, counter)
      If counter > 1 Then unfinished = True
      Exit Function
   End If
Next j
Next i

End Function


Function nState(vertex As Range, state As Integer) As Integer
If vertex.Offset(-1, 0).Value = state Then nState = nState + 1
If vertex.Offset(1, 0).Value = state Then nState = nState + 1
If vertex.Offset(0, -1).Value = state Then nState = nState + 1
If vertex.Offset(0, 1).Value = state Then nState = nState + 1
End Function


Function trace(ByVal vertex As Range, pointer As Integer, countdown As Integer, origin As String, counter As Integer) As Range
If vertex.Address = origin Or countdown = 0 Then
   Set trace = vertex
ElseIf pointer = 1 Then
   If vertex.Offset(-1, 0).Value = 5 Then
      counter = counter - 1
      Set trace = trace(vertex.Offset(-2, 0), 4, 3, origin, counter)
   Else
      Set trace = trace(vertex, 2, countdown - 1, origin, counter)
   End If
ElseIf pointer = 2 Then
   If vertex.Offset(0, 1).Value = 5 Then
      counter = counter - 1
      Set trace = trace(vertex.Offset(0, 2), 1, 3, origin, counter)
   Else
      Set trace = trace(vertex, 3, countdown - 1, origin, counter)
   End If
ElseIf pointer = 3 Then
   If vertex.Offset(1, 0).Value = 5 Then
      counter = counter - 1
      Set trace = trace(vertex.Offset(2, 0), 2, 3, origin, counter)
   Else
      Set trace = trace(vertex, 4, countdown - 1, origin, counter)
   End If
ElseIf pointer = 4 Then
   If vertex.Offset(0, -1).Value = 5 Then
      counter = counter - 1
      Set trace = trace(vertex.Offset(0, -2), 3, 3, origin, counter)
   Else
      Set trace = trace(vertex, 1, countdown - 1, origin, counter)
   End If
End If
End Function


Sub clearlines()
Call redoFormatting
Sheets("Sheet2").Range("B22:R38").Copy Destination:=Range("BD20")
End Sub

# VBA
Smth
Dim i As Integer
Dim j As Integer
For i = 1 To 3
For j = 1 To 3
mat(i, j) = Cells(i, j)
Next
Next
Dim n As Integer
Dim m As Integer
Dim d As Integer
For n = 1 To 3
d = 1
Do While (d = 1)
For m = 1 To 2
Dim s As Integer
If mat(m, n) < mat(m + 1, n) Then
s = mat(m, n): mat(m, n) = mat(m + 1, n): mat(m + 1, n) = s: d = 1
Else: d = 0
End If
Next
Loop
Next
Dim prompt As String
Dim i1 As Integer
Dim j1 As Integer
prompt = ""
For i1 = 1 To UBound(mat, 1)
  For j1 = 1 To UBound(mat, 2)
   prompt = prompt & mat(i1, j1) & " "
  Next j1
  prompt = prompt & Chr(13)
Next i1
MsgBox prompt
Range("E1:G3") = mat
End Sub

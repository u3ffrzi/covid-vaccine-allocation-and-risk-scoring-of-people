Attribute VB_Name = "Module1"
Public classValue() As Variant
Sub f()
rno = Range("b2").CurrentRegion.Rows.Count + 1
For i = 1 To Range("j1").CurrentRegion.Rows.Count
j = 1

While j <= 1
Range("c" + CStr(rno)).Value = Range("j" + CStr(i)).Value
If j = 1 Then
Range("d" + CStr(rno)).Value = "S"
ElseIf j = 2 Then

Range("d" + CStr(rno)).Value = "BP"
ElseIf j = 3 Then

Range("d" + CStr(rno)).Value = "CP"
End If
rno = rno + 1

j = j + 1

Wend

Next

End Sub
Sub defineName()

p = 0

For i = 2 To Range("c1").CurrentRegion.Rows.Count
If Range("c" + CStr(i)).Value <> Range("c" + CStr(i - 1)).Value Then
d = 1
p = p + 1
Range("f" + CStr(p)).Value = Range("c" + CStr(i)).Value
Cells(p, 6 + d).Value = Range("d" + CStr(i)).Value
d = d + 1
Else

Cells(p, 6 + d).Value = Range("d" + CStr(i)).Value
d = d + 1
End If

Next







End Sub

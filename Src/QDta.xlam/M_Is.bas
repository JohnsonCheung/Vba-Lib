Attribute VB_Name = "M_Is"
Function IsSimTyLvs(A$) As Boolean
Dim Ay$(): Ay = SslSy(A)
If AyIsEmp(Ay) Then Exit Function
Dim I
For Each I In Ay
    Stop '
'   If Not IsSimTyStr(Ay) Then Exit Function
Next
IsSimTyLvs = True
End Function



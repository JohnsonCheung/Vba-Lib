Attribute VB_Name = "M_Is"
Option Explicit

Function IsSimTyLvs(A$) As Boolean
Dim Ay$(): Ay = SslSy(A)
If AyIsEmp(Ay) Then Exit Function
Dim I
For Each I In Ay
   If Not IsSimTyStr(Ay) Then Exit Function
Next
IsSimTyLvs = True
End Function

Function IsSimTySsl(A) As Boolean
Dim Ay$(): Ay = SslSy(A)
If AyIsEmp(Ay) Then Exit Function
Dim I
For Each I In Ay
    If Not IsSimTyStr(Ay) Then Exit Function
Next
IsSimTySsl = True
End Function

Function IsSimTyStr(A) As Boolean
Select Case UCase(A)
Case "TXT", "NBR", "LGC", "DTE", "OTH": IsSimTyStr = True
End Select
End Function

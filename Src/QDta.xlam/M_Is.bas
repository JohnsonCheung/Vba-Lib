Attribute VB_Name = "M_Is"
Option Explicit

Property Get IsSimTySsl(A) As Boolean
Dim Ay$(): Ay = SslSy(A)
If AyIsEmp(Ay) Then Exit Property
Dim I
For Each I In Ay
    If Not IsSimTyStr(Ay) Then Exit Function
Next
IsSimTySsl = True
End Property

Property Get IsSimTyStr(A) As Boolean
Select Case UCase(A)
Case "TXT", "NBR", "LGC", "DTE", "OTH": IsSimTyStr = True
End Select
End Property

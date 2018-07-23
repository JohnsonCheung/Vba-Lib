Attribute VB_Name = "M_Ws"
Option Explicit

Function Ws(Optional Hid As Boolean) As Worksheet
Dim O As Worksheet: Set O = NewWs(Vis:=Not Hid)
Stop '
'WsA1(O).Value = "*Ds " & A.DsNm
Dim At As Range, J%
Set At = WsRC(O, 2, 1)
Stop '
'For J = 0 To DsNDt(A)
'    Set At = DtAt(A.DtAy(J), At, J)
'Next
Set Ws = O
End Function

Attribute VB_Name = "G_Ay"
Option Explicit

Function Pop(Ay)
Pop = AyLasEle(Ay)
AyRmvLasNEle Ay
End Function

Sub ReSz(Ay, U&)
If U < 0 Then
    Erase Ay
Else
    ReDim Preserve Ay(U)
End If
End Sub

Function Sz&(Ay)
On Error Resume Next
Sz = UBound(Ay) + 1
End Function

Function UB&(Ay)
UB = Sz(Ay) - 1
End Function

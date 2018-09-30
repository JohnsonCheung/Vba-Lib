Attribute VB_Name = "VbAyDic"
Option Explicit
Private Sub Z_Aydic_to_KeyCntMulItmCol_Dry()
Dim A As New Dictionary, Act()
A.Add "A", Array(1, 2, 3)
A.Add "B", Array(2, 3, 4)
A.Add "C", Array()
A.Add "D", Array("X")
Act = Aydic_to_KeyCntMulItmColDry(A)
Ass Sz(Act) = 4
Ass IsEqAy(Act(0), Array("A", 3, 1, 2, 3))
Ass IsEqAy(Act(1), Array("B", 3, 2, 3, 4))
Ass IsEqAy(Act(2), Array("C", 0))
Ass IsEqAy(Act(3), Array("D", 1, "X"))
End Sub
Function IsAydic(A As Dictionary) As Boolean
If Not IsAyOfStr(A.Keys) Then Exit Function
If Not IsAyOfAy(A.Items) Then Exit Function
IsAydic = True
End Function
Function Aydic_to_KeyCntMulItmColDry(A As Dictionary) As Variant()
If A.Count = 0 Then Exit Function
Dim O(), K, Dr(), Ay, J&
ReDim O(A.Count - 1)
For Each K In A.Keys
    Ay = A(K): If Not IsArray(Ay) Then Stop
    O(J) = AyIns2(Ay, K, Sz(Ay))
    J = J + 1
Next
Aydic_to_KeyCntMulItmColDry = O
End Function

Attribute VB_Name = "M_DicAy"
Option Explicit

Property Get DicAy_Dr(DicAy, K) As Variant()
Dim U%: U = UB(DicAy)
Dim O()
ReDim O(U + 1)
Dim I, Dic As Dictionary, J%
J = 1
O(0) = K
For Each I In DicAy
   Set Dic = I
   If Dic.Exists(K) Then O(J) = Dic(K)
   J = J + 1
Next
DicAy_Dr = O
End Property

Property Get DicAy_Ky(DicAy) As Variant()
Dim O(), Dic As Dictionary, I
For Each I In DicAy
   Set Dic = I
   PushNoDupAy O, Dic.Keys
Next
DicAy_Ky = O
End Property

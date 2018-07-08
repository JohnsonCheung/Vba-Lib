Attribute VB_Name = "M_VblAy"
Option Explicit
Property Get VblLy_Dry(VblLy$()) As Variant()
If AyIsEmp(VblLy) Then Exit Property
Dim O()
   Dim I
   For Each I In VblLy
       Push O, AyTrim(SplitVBar(CStr(I)))
   Next
VblLy_Dry = O
End Property



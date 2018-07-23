Attribute VB_Name = "M_DotNy"
Option Explicit

Function DotNy_Dry(DotNy$()) As Variant()
If AyIsEmp(DotNy) Then Exit Function
Dim O(), I
For Each I In DotNy
   With Brk1(I, ".")
       Push O, ApSy(.S1, .S2)
   End With
Next
DotNy_Dry = O
End Function

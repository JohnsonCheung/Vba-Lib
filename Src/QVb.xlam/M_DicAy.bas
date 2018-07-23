Attribute VB_Name = "M_DicAy"
Option Explicit

Function DicAy_Ky(A) As Variant()
Dim O(), I
For Each I In A
   PushNoDupAy O, CvDic(I).Keys
Next
DicAy_Ky = O
End Function

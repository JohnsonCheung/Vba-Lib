Attribute VB_Name = "M_IxAy"
Option Explicit
Property Get IxAy_IsAllGE0(IxAy&()) As Boolean
Dim J&
For J = 0 To UB(IxAy)
    If IxAy(J) = -1 Then Exit Property
Next
IxAy_IsAllGE0 = True
End Property
Property Get IxAy_IsParitial_of_0toU(IxAy, U&) As Boolean
Const CSub$ = "Ass IxAy_IsParitial_of_0toU"
Const Msg$ = "{IxAy} is not PartialIx-of-{U}." & _
"|PartialIxAy-Of-U is defined as:" & _
"|It should be Lng()" & _
"|It should have 0 to U elements" & _
"|It should have each element of value between 0 and U" & _
"|It should have no dup element" & _
"|All elements should have value equal or less than U"

If Not IsLngAy(IxAy) Then Exit Property
If AyIsEmp(IxAy) Then IxAy_IsParitial_of_0toU = True: Exit Property
If AyHasDupEle(IxAy) Then Exit Property
Dim I
For Each I In IxAy
   If 0 > I Or I > U Then Exit Property
Next
IxAy_IsParitial_of_0toU = True
End Property


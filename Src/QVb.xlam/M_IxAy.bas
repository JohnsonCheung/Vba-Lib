Attribute VB_Name = "M_IxAy"
Option Explicit

Function IxAy_IsAllGE0(A) As Boolean
Dim J&
For J = 0 To UB(A)
    If A(J) = -1 Then Exit Function
Next
IxAy_IsAllGE0 = True
End Function

Function IxAy_IsParitial_of_0toU(A, U&) As Boolean
Const CSub$ = "Ass IxAy_IsParitial_of_0toU"
Const Msg$ = "{IxAy} is not PartialIx-of-{U}." & _
"|PartialIxAy-Of-U is defined as:" & _
"|It should be Lng()" & _
"|It should have 0 to U elements" & _
"|It should have each element of value between 0 and U" & _
"|It should have no dup element" & _
"|All elements should have value equal or less than U"

If Not IsLngAy(A) Then Exit Function
If AyIsEmp(A) Then IxAy_IsParitial_of_0toU = True: Exit Function
If AyHasDupEle(A) Then Exit Function
Dim I
For Each I In A
   If 0 > I Or I > U Then Exit Function
Next
IxAy_IsParitial_of_0toU = True
End Function
Private Sub ZZ_IxAy_IsParitial_of_0toU()
Ass IxAy_IsParitial_of_0toU(ApLngAy(1, 2, 3, 4), 4) = True
Ass IxAy_IsParitial_of_0toU(ApLngAy(0, 1, 2, 3, 4), 4) = True
Ass IxAy_IsParitial_of_0toU(ApLngAy(1, 1, 3, 4), 4) = False
Ass IxAy_IsParitial_of_0toU(ApLngAy(5, 3, 4), 4) = False
End Sub

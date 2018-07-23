Attribute VB_Name = "M_Ay"
Option Explicit

Function AyC0Dry(Constant, Ay) As Variant()
If AyIsEmp(Ay) Then Exit Function
Dim O(), I
For Each I In Ay
   Push O, Array(Constant, I)
Next
AyC0Dry = O
End Function

Function AyC1Dry(Ay, Constant) As Variant()
If AyIsEmp(Ay) Then Exit Function
Dim O(), I
For Each I In Ay
   Push O, Array(I, Constant)
Next
AyC1Dry = O
End Function

Function AyDt(A, Optional FldNm$ = "Itm", Optional DtNm$ = "Ay") As Dt
Dim O(), J&
For J = 0 To UB(A)
    Push O, Array(A(J))
Next
AyDt = Dt(DtNm, FldNm, O)
End Function

Attribute VB_Name = "M_Ay"
Option Explicit
Property Get AyDt(A, Optional FldNm$ = "Itm", Optional DtNm$ = "Ay") As Dt
Dim O As Dt
O.DtNm = DtNm
O.Fny = ApSy(FldNm)
Dim ODry(), J%
For J = 0 To UB(A)
    Push ODry, Array(A(J))
Next
O.Dry = ODry
AyDt = O
End Property


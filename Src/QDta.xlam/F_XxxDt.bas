Attribute VB_Name = "F_XxxDt"
Option Explicit
Property Get AyDt(A, Optional FldNm$ = "Itm", Optional DtNm$ = "Ay") As Dt
Dim ODry(), J&
For J = 0 To UB(A)
    Push ODry, Array(A(J))
Next
AyDt = Dt(DtNm, FldNm, ODry)
End Property

Property Get DicDt(A As Dictionary, Optional DtNm$ = "Dic", Optional InclDicValTy As Boolean) As Dt
Dim Dry()
Dry = DicDry(A, InclDicValTy)
Dim F$
    If InclDicValTy Then
        F = "Key Val Ty"
    Else
        F = "Key Val"
    End If
Set DicDt = Dt(DtNm, F, Dry)
End Property



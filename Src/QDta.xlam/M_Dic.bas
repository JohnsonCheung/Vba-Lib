Attribute VB_Name = "M_Dic"
Option Explicit
Property Get DicDt(A As Dictionary, Optional DtNm$ = "Dic", Optional InclDicValTy As Boolean) As Dt
Me.DtNm = DtNm
B_Dry = DicDry(A, InclDicValTy)
Dim F$
    If InclDicValTy Then
        F = "Key Val Ty"
    Else
        F = "Key Val"
    End If
B_Fny = SslSy(F)
Set InitByDic = Me
End Property


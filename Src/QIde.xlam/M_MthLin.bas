Attribute VB_Name = "M_MthLin"
Option Explicit
Function MthLin_EnsPrivate(A) As StrOpt
Dim P As Parse: P = ParseKwMdy(NewParse(A))
If Not P.IsOk Then Exit Function
Dim P1 As Parse: P1 = ParseKwMthTy(P)
If Not P.IsOk Then Exit Function
If P.Er_or_Ok = "Private" Then MthLin_EnsPrivate = StrOpt(A): Exit Function
MthLin_EnsPrivate = StrOpt("Private " & P.Lin)
End Function
Function MthLin_Key$(A)
With SrcLin_MthBrk(A)
    MthLin_Key = FmtQQ("?:?:?", .Mdy, .Ty, .MthNm)
End With
End Function

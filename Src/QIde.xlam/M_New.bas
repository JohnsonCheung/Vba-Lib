Attribute VB_Name = "M_New"
Option Explicit
Function NewMthPrm(MthPrmStr$) As MthPrm
Stop
'Dim L$: L = MthPrmStr
'Dim TyChr$
'With MthPrm
'    .IsOpt = ParseHasPfxSpc(L, "Optional")
'    .IsPrmAy = ParseHasPfxSpc(L, "ParamArray")
'    .Nm = ParseNm(L)
'    .Ty.TyChr = ParseOneOfChr(L, "!@#$%^&")
'    .Ty.IsAy = ParseHasPfx(L, "()")
'End With
End Function
Function NewMthSrc(Nm$, Ly$()) As SrcItm
NewMthSrc.Nm = Nm
NewMthSrc.Ly = Ly
NewMthSrc.SrcTy = eMth
End Function
Function NewSrcItmCnt(N%, NPub%, NPrv%) As SrcItmCnt
With NewSrcItmCnt
    .N = N
    .NPrv = NPrv
    .NPub = NPub
End With
End Function
Function NewTySrc(Nm$, Ly$()) As SrcItm
NewTySrc.Nm = Nm
NewTySrc.Ly = Ly
NewTySrc.SrcTy = eDtaTy
End Function

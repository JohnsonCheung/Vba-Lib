Attribute VB_Name = "M_DrsLy"
Option Explicit

Function DrsLy_Drs(DrsLy$()) As Drs
Dim Fny$(): Fny = SslSy(DrsLy(0))
Dim J&, Dry()
If IsSimTySsl(DrsLy(2)) Then
    Dim TyAy() As eSimTy
    For J = 3 To UB(DrsLy)
        Push Dry, SslDr(DrsLy(J), TyAy)
    Next
Else
    For J = 2 To UB(DrsLy)
        Push Dry, SslSy(DrsLy(J))
    Next
End If
Set DrsLy_Drs = Drs(Fny, Dry)
End Function

Function DrsLy_InsBrkLin(TblLy$(), ColNm$) As String()
Dim Hdr$: Hdr = TblLy(1)
Dim Fny$():
    Fny = SplitVBar(Hdr)
    Fny = AyRmvFstEle(Fny)
    Fny = AyRmvLasEle(Fny)
    Fny = AyTrim(Fny)
Dim Ix%
    Ix = AyIx(Fny, ColNm)
Dim DryLy$()
    DryLy = AyWhExclAtCnt(TblLy, 0, 2)
Dim O$()
    Push O, TblLy(0)
    Push O, TblLy(1)
    PushAy O, DryLy_InsBrkLin(DryLy, Ix)
DrsLy_InsBrkLin = O
End Function

Private Sub ZZ_DrsLy_InsBrkLin()
Dim TblLy$()
Dim Act$()
Dim Exp$()
TblLy = FtLy(TstResPth & "DrsLy_InsBrkLin.txt")
Act = DrsLy_InsBrkLin(TblLy, "Tbl")
Exp = FtLy(TstResPth & "DrsLy_InsBrkLin_Exp.txt")
'AyPair_EqChk Exp, Act
End Sub

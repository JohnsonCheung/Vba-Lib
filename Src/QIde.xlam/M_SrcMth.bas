Attribute VB_Name = "M_SrcMth"
Option Explicit
Function SrcMth_BdyLines$(A$(), MthNm$)
SrcMth_BdyLines = JnCrLf(SrcMth_BdyLy(A, MthNm))
End Function
Function SrcMth_BdyLy(A$(), MthNm$) As String()
Dim FmTo() As FmTo: FmTo = SrcMth_FmToAy(A, MthNm)
Dim O$(), J%
For J = 0 To UB(FmTo)
   PushAy O, AyWhFmTo(A, FmTo(J))
Next
SrcMth_BdyLy = O
End Function
Function SrcMth_FmToAy(A$(), MthNm$) As FmTo()
Dim IxAy&(), O() As FmTo, M As FmTo
IxAy = SrcMth_LxAy(A, MthNm)
Dim J%
For J = 0 To UB(IxAy)
   M.FmIx = IxAy(J)
   M.ToIx = SrcMthLx_ToLx(A, M.FmIx)
   Push O, M
Next
SrcMth_FmToAy = O
End Function
Function SrcMth_Lno%(A$(), MthNm, Optional PrpTy$)
If AyIsEmp(A) Then Exit Function
If PrpTy <> "" Then
   If Not AyHas(Array("Get Let Set"), PrpTy) Then Stop
End If
Dim FunTy$: FunTy = "Property " & PrpTy
Dim Lno&
Lno = 0
Const IMthNm% = 2
Dim M As MthBrk
Dim Lin
For Each Lin In A
   Lno = Lno + 1
   M = SrcLin_MthBrk(Lin)
   If M.MthNm = "" Then GoTo Nxt
   If M.MthNm <> MthNm Then GoTo Nxt
   If PrpTy <> "" Then
       If M.Ty <> FunTy Then GoTo Nxt
   End If
   SrcMth_Lno = Lno
   Exit Function
Nxt:
Next
SrcMth_Lno = 0
End Function
Function SrcMth_LnoCnt(A$(), MthNm$) As LnoCnt
End Function
Function SrcMth_LnoCntAy(A$(), MthNm$) As LnoCnt()
Dim FmAy&(): FmAy = SrcMth_LxAy(A, MthNm)
Dim O() As LnoCnt, J%
Dim ToIx&
Dim FT As FmTo
Dim LnoCnt As LnoCnt
For J = 0 To UB(FmAy)
   ToIx = SrcMthLx_ToLx(A, FmAy(J))
   Set FT = FmTo(FmAy(J), ToIx)
   LnoCnt = FmTo_LnoCnt(FT)
   Push O, LnoCnt
Next
SrcMth_LnoCntAy = O
End Function
Function SrcMth_Lx%(A$(), MthNm$, Optional Fm&)
Dim I%, Nm$
For I = Fm To UB(A)
    Nm = SrcLin_MthNm(A(I))
    If Nm = MthNm Then
        SrcMth_Lx% = I
        Exit Function
    End If
Next
SrcMth_Lx = -1
End Function
Function SrcMth_LxAy(A$(), MthNm$) As Long()
Dim Ix&
   Ix = SrcMth_Lx(A, MthNm)
   If Ix = -1 Then Exit Function

Dim O&()
   Push O, Ix
   If Not HasPfx(SrcLin_MthTy(A(Ix)), "Property") Then
       SrcMth_LxAy = O
       Exit Function
   End If

   Dim J%, Fm&
   For J = 1 To 2
       Fm = Ix + 1
       Ix = SrcMth_Lx(A, MthNm, Fm)
       If Ix = -1 Then
           SrcMth_LxAy = O
           Exit Function
       End If
       Push O, Ix
   Next
SrcMth_LxAy = O
End Function
Function SrcMth_RRCC(A$(), MthNm$) As RRCC
Dim R&, C&, Ix&
Ix = SrcMth_Lx(A, MthNm)
R = Ix + 1
C = SrcLin_MthNmPos(A(Ix))
SrcMth_RRCC = RRCC(R, R, C + 1, C + Len(MthNm))
End Function
Private Sub ZZ_SrcMth_BdyLy()
Dim Src$(): Src = ZZSrc
Dim MthNm$: MthNm = "A"
Dim Act$()
Act = SrcMth_BdyLy(Src, MthNm)
End Sub

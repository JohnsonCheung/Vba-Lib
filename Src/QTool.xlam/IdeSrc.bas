Attribute VB_Name = "IdeSrc"
Option Explicit
Function SrcMthKy(A$(), Optional PjNm$ = "Pj", Optional MdNm$ = "Md", Optional IsWrap As Boolean) As String()
Dim L$(): L = SrcMthLinAy(A) ' MthLinAy
SrcMthKy = AyMapAvSy(L, "LinMthKey", Array(PjNm, MdNm, IsWrap))
End Function
Function SrcMthLinDry(A$()) As Variant()
Dim L
For Each L In AyNz(A)
    PushNonBlankAy SrcMthLinDry, LinMthDr(L)
Next
End Function
Function SrcMthLinDryWP(A$()) As Variant()
Dim L
For Each L In AyNz(A)
    PushISomSz SrcMthLinDryWP, LinMthDrWP(L)
Next
End Function
Function SrcMthIxDr(A$(), MthIx) As Variant()
Dim M$, T$, N$, Lines$, C%, MthTo%
MthTo = SrcMthIxTo(A, MthIx)
C = MthTo - MthIx + 1
Lines = JnCrLf(AyMid(A, MthIx, C))
LinMthBrkAsg A(MthIx), M, T, N
SrcMthIxDr = Array(M, T, N, Lines)
End Function
Function SrcMthIxDrWh(A$(), MthIx, B As WhMth, Optional C As MthBrkOpt) As Variant()
SrcMthIxDrWh = Array(MthIx, A(MthIx))
Exit Function
Dim F% ' MthFmno
Dim Mth
Dim MthTo%
MthTo = SrcMthIxTo(A, MthIx)
C = MthTo - MthIx + 1
Mth = JnCrLf(AyWhFmTo(A, MthIx, MthTo))
Mth = SrcContLin(A, F)
SrcMthIxDrWh = Array(F, C, Mth)
End Function
Function SrcMthDry(A$()) As Variant()
Dim Ix
For Each Ix In AyNz(SrcMthIx(A))
    Push SrcMthDry, SrcMthIxDr(A, Ix)
Next
End Function
Function SrcMthDryWh(A$(), B As WhMth, Optional C As MthWhOpt) As Variant()
Dim Ix
For Each Ix In AyNz(SrcMthIxWh(A, C))
    If SrcMthIxWh(A, Ix, C) Then
        Push SrcMthDryWh, SrcMthIxDrWh(A, Ix, B, C)
    End If
Next
End Function
Function SrcMthLinAy(A$(), Optional WhMdyAy, Optional WhKdAy) As String()
Dim O$(), L
For Each L In AyNz(SrcMthIx(A))
    Push O, SrcContLin(A, L)
Next
SrcMthLinAy = O
End Function
Function SrcHasMth(A$(), MthNm) As Boolean
Dim L
For Each L In AyNz(A)
    If LinMthNm(L) = MthNm Then SrcHasMth = True: Exit Function
Next
End Function
Function SrcMthIxWh(A$(), MthIx, Optional B As MthWhOpt) As Boolean
SrcMthIxWh = True
Dim M$(), K$(), WhMdy As Boolean, SelKd As Boolean
'M = CvWhMdy(WhMdy): WhMdy = Sz(M) > 0
'K = CvWhMthKd(WhKdAy): SelKd = Sz(K) > 0
End Function
Function SrcMthIx(A$()) As Integer()
Dim O%(), J%
For J = 0 To UB(A)
    If IsMthLin(A(J)) Then
        Push O, J
    End If
Next
SrcMthIx = O
End Function
Function SrcMthFT(A$()) As FTIx()
Dim F%(): F = SrcMthIx(A)
Dim U%: U = UB(F)
If U = -1 Then Exit Function
Dim O() As FTIx
ReDim O(U)
Dim J%
For J = 0 To U
    Set O(J) = FTIx(F(J), SrcMthIxTo(A, F(J)))
Next
SrcMthFT = O
End Function
Function SrcContLin$(A$(), Ix)
Dim O$(), J%, L$
For J = Ix To UB(A)
    L = A(J)
    If Right(L, 2) <> " _" Then
        Push O, L
        SrcContLin = Join(O, "")
        Exit Function
    End If
    Push O, RmvLasNChr(L, 2)
Next
ErImposs
End Function
Function SrcDclLinCnt%(A$())
Dim I&
    I = SrcFstMthIx(A)
    If I = -1 Then
        SrcDclLinCnt = Sz(A)
        Exit Function
    End If
    I = SrcMthIxRmkFm(A, I)
Dim O&, L$
    For I = I - 1 To 0 Step -1
        If IsCdLin(A(I)) Then
            O = I + 1
            GoTo X
        End If
    Next
X:
SrcDclLinCnt = O
End Function
Function SrcDclLines$(A$())
SrcDclLines = Join(SrcDclLy(A), vbCrLf)
End Function
Function SrcDclLy(A$()) As String()
If Sz(A) = 0 Then Exit Function
Dim N&
   N = SrcDclLinCnt(A)
If N <= 0 Then Exit Function
SrcDclLy = AyFstNEle(A, N)
End Function
Function SrcMthKeyLinesDic(A$(), Optional PjNm$, Optional MdNm$, Optional ExlDcl As Boolean) As Dictionary
Dim L%(): L = SrcMthIx(A)
Dim K$
Dim O As New Dictionary
    If Not ExlDcl Then
        If PjNm = "" And MdNm = "" Then
            K$ = PjNm & "." & MdNm & ".*Dcl"
        Else
            K = "*Dcl"
        End If
        O.Add K, SrcDclLines(A)
    End If
    If Sz(L) = 0 Then GoTo X
    Dim MthNm$, Lin$, Lines$, Lx
    For Each Lx In L
        Lin = SrcContLin(A, CInt(Lx))
        MthNm = LinMthNm(Lin):               If MthNm = "" Then Stop
        Lines = SrcMthIxLines(A, Lx): If Lines = "" Then Stop
        K = LinMthKey(Lin, PjNm, MdNm)
        O.Add K, Lines
    Next
X:
Set SrcMthKeyLinesDic = O
End Function
Function SrcFstMthIx&(A$())
Dim J%
For J = 0 To UB(A)
   If IsMthLin(A(J)) Then
       SrcFstMthIx = J
       Exit Function
   End If
Next
SrcFstMthIx = -1
End Function
Function SrcMthNmIx(A$(), MthNm, Optional Fm% = 0) As Integer()
Dim J%, O%()
For J = Fm To UB(A)
    If LinMthNm(A(J)) = MthNm Then
        If IsPrpLin(A(J)) Then
            SrcMthNmIx = ItmAddAy(J, SrcMthNmIx(A, MthNm, J + 1))
        Else
            SrcMthNmIx = ApIntAy(J)
        End If
        Exit Function
    End If
Next
End Function
Function SrcMthNmFC(A$(), MthNm) As FmCnt()
SrcMthNmFC = FTIxAyFC(SrcMthNmFT(A, MthNm))
End Function
Function SrcMthNmFT(A$(), MthNm) As FTIx()
Dim Toix%, Fm
For Each Fm In AyNz(SrcMthNmIx(A, MthNm))
    Toix = SrcMthIxTo(A, Fm)
    Push SrcMthNmFT, FTIx(Fm, Toix)
Next
End Function
Function SrcMthRmkFC(A$(), MthNm) As FmCnt()
Dim Fmix%, Toix%, Fm
For Each Fm In AyNz(SrcMthNmIx(A, MthNm))
    Fmix = SrcMthIxRmkFm(A, Fm)
    Toix = SrcMthIxTo(A, Fm)
    PushObj SrcMthRmkFC, FmCnt(Fmix + 1, Toix - Fmix + 1)
Next
End Function
Function SrcMthFTAy(A$(), MthNm) As FTIx()
Dim F%()
F = SrcMthNmIx(A, MthNm): If Sz(F) <= 0 Then Exit Function
Dim O() As FTIx
ReDim O(UB(F))
Dim J%
For J = 0 To UB(F)
    Set O(J) = FTIx(F(J), SrcMthIxTo(A, F(J)))
Next
SrcMthFTAy = O
End Function
Function SrcMthLin$(A$(), MthNm)
SrcMthLin = SrcContLin(A, SrcMthNmIx(A, MthNm))
End Function
Function SrcMthIxLinesWithRmk$(A$(), MthIx)
SrcMthIxLinesWithRmk = ApLines(SrcMthIxRmk(A, MthIx), SrcMthIxLines(A, MthIx))
End Function
Function SrcMthIxLines$(A$(), MthIx)
Dim L2%
L2 = SrcMthIxTo(A, MthIx): If L2 = 0 Then Stop
SrcMthIxLines = Join(AyWhFmTo(A, MthIx, L2), vbCrLf)
End Function
Function SrcMthNy(A$()) As String()
Dim L
For Each L In AyNz(SrcMthIx(A))
    PushNonBlankStr SrcMthNy, LinMthNm(A(L))
Next
End Function
Function SrcMthNyWh(A$(), B As WhMth) As String()
Dim L, O$()
For Each L In AyNz(SrcMthIx(A))
    PushNonBlankStr O, LinMthNmWh(A(L), B)
Next
SrcMthNyWh = AyWhDist(O)
End Function
Function SrcMthFNy(A$()) As String()
Dim L, O$()
For Each L In AyNz(SrcMthIx(A))
    PushNonBlankStr O, LinMthFNm(A(L))
Next
End Function
Function SrcMthFNyWh(A$(), B As WhMth) As String()
Dim L
For Each L In AyNz(SrcMthIx(A))
    PushNonBlankStr SrcMthFNyWh, LinMthFNmWh(A(L), B)
Next
End Function
Function SrcMthIxRmkFm%(A$(), MthIx)
Dim M1&
    Dim J&
    For J = MthIx - 1 To 0 Step -1
        If IsCdLin(A(J)) Then
            M1 = J
            GoTo M1IsFnd
        End If
    Next
    M1 = -1
M1IsFnd:
Dim M2&
    For J = M1 + 1 To MthIx - 1
        If Trim(A(J)) <> "" Then
            M2 = J
            GoTo M2IsFnd
        End If
    Next
    M2 = MthIx
M2IsFnd:
SrcMthIxRmkFm = M2
End Function
Function SrcMthIxTo%(A$(), MthIx)
Dim T$, F$, J%
T = LinMthKd(A(MthIx)): If T = "" Then Stop
F = "End " & T
For J = MthIx + 1 To UB(A)
    If IsPfx(A(J), F) Then SrcMthIxTo = J: Exit Function
Next
Stop
End Function
Function SrcMth10Dry(A$()) As Variant()
'Mdy ShtTy Nm Sfx Prm Ret ShtRmk Lno Cnt Lines TopRmk
'1   2     3  4   5   6   7      8   9   10    11
Dim L, O(), Lin$, M7$(), Ay%(), Lines$
Ay = SrcMthIx(A)
If Sz(Ay) = 0 Then Exit Function
For Each L In Ay
    Lin = SrcContLin(A, CStr(L))
    M7 = LinMth7Dr(Lin)
    If Sz(M7) = 0 Then GoTo X
    Lines = SrcMthIxLines(A, L)
    PushAy M7, Array(L + 1, LinCnt(Lines), Lines)
    Push O, M7
X:
Next
SrcMth10Dry = O
End Function
Function SrcMth7Dry(A$()) As Variant()
Dim L, O()
For Each L In AyNz(SrcMthLinAy(A))
    PushWithSz O, LinMth7Dr(L)
Next
SrcMth7Dry = O
End Function
Function SrcMthDot(A$(), Optional MthRe As RegExp, Optional MthExlAy$, Optional WhMdyAy, Optional WhKdAy) As String()
Stop '
'SrcMthDot = AyMapSy(SrcMthBrk(A, MthPatn, MthExlAy, WhMdyA, WhKdAy), "MthBrkDot")
End Function
Function SrcMthLinesWithRmk$(A$(), MthNm)
Dim L%(): L = SrcMthNmIx(A, MthNm)
If Sz(L) = 0 Then Exit Function
Dim MthIx, O$()
For Each MthIx In L
    Push O, SrcMthIxLinesWithRmk(A, MthIx)
Next
SrcMthLinesWithRmk = Join(O, vbCrLf & vbCrLf)
End Function
Function SrcMthIxRmk$(A$(), MthIx)
Dim O$(), J%, L$, I%
Dim Fm%: Fm = SrcMthIxRmkFm(A, MthIx)

For J = Fm To MthIx - 1
    If IsRmkLin(A(J)) Then Push O, L
Next
SrcMthIxRmk = Join(O, vbCrLf)
End Function
Function SrcMthLines$(A$(), MthNm$)
Dim I, O$()
For Each I In AyNz(SrcMthNmIx(A, MthNm))
    Push O, SrcMthIxLines(A, I)
Next
SrcMthLines = Join(O, vbCrLf & vbCrLf)
End Function
Function SrcMthLinesDic(A$(), Optional ExlDcl As Boolean) As Dictionary
Dim L%(): L = SrcMthIx(A)
Dim O As New Dictionary
    If Not ExlDcl Then O.Add "*Dcl", SrcDclLines(A)
    If Sz(L) = 0 Then GoTo X
    Dim MthNm$, Lin$, Lines$, Lx
    For Each Lx In L
        Lin = A(Lx)
        MthNm = LinMthNm(Lin):            If MthNm = "" Then Stop
        Lines = SrcMthIxLines(A, Lx): If Lines = "" Then Stop
        If O.Exists(MthNm) Then
            If Not IsPrpLin(Lin) Then Stop
            O(MthNm) = O(MthNm) & vbCrLf & vbCrLf & Lines
        Else
            O.Add MthNm, Lines
        End If
    Next
X:
Set SrcMthLinesDic = O
End Function
Function SrcMthDic(A$(), Optional PjNm$, Optional MdNm$) As Dictionary
Dim Ix, Key$, Lines$, IxAy%(), O As New Dictionary
IxAy = SrcMthIx(A)
For Each Ix In AyNz(IxAy)
    Key = LinMthKey(A(Ix), PjNm, MdNm)
    Lines = SrcMthIxLines(A, Ix)
    If O.Exists(Key) Then
        If Not IsPrpLin(A(Ix)) Then Stop
        O(Key) = O(Key) & vbCrLf & Lines
    Else
        O.Add Key, Lines
    End If
Next
Set SrcMthDic = O
End Function
Function SrcMthBrkWh(A$(), B As WhMth) As Variant()
'Dim L, O(), B$(), Re As RegExp, ExlLikAy$()
'If MthPatn <> "." Then Set Re = New RegExp: Re.Pattern = MthPatn
'ExlLikAy = SslSy(MthExlAy)
'For Each L In AyNz(SrcMthIx(A)) ', WhMdyA, WhKdAy))
'    B = LinMthBrk(A(L))
'    If IsNmSel(B(2), Re, ExlLikAy) Then
'        Push O, LinMthBrk(A(L))
'    End If
'Next
'SrcMthBrk = O
End Function
Function SrcMthBrk(A$()) As Variant()
Dim L
For Each L In AyNz(SrcMthIx(A)) ', WhMdyA, WhKdAy))
    Push SrcMthBrk, LinMthBrk(L)
Next
End Function

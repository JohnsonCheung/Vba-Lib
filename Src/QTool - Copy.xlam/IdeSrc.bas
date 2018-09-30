Attribute VB_Name = "IdeSrc"
Option Explicit
Function SrcMthCxtFT(A$(), MthNm$) As FTNo()
Dim P() As FTIx
Dim Ix() As FTIx: Ix = SrcMthNmFT(A, MthNm)
SrcMthCxtFT = AyMapPXInto(Ix, "CxtIx", A, P)
End Function
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

Function SrcMthIxDr(A$(), MthIx, Optional B As WhMth, Optional C As MthBrkOpt) As Variant()
'Dim M$, T$, N$, Lines$, C%, MthTo%
'MthTo = SrcMthIxTo(A, MthIx)
'C = MthTo - MthIx + 1
'Lines = JnCrLf(AyMid(A, MthIx, C))
'LinMthBrkAsg A(MthIx), M, T, N
'SrcMthIxDr = Array(M, T, N, Lines)

SrcMthIxDr = Array(MthIx, A(MthIx))
Exit Function
Dim F% ' MthFmno
Dim Mth
Dim MthTo%
MthTo = SrcMthIxTo(A, MthIx)
C = MthTo - MthIx + 1
Mth = JnCrLf(AyWhFmTo(A, MthIx, MthTo))
Mth = SrcContLin(A, F)
SrcMthIxDr = Array(F, C, Mth)
End Function

Function SrcMthDry(A$(), Optional B As WhMth, Optional C As MthBrkOpt) As Variant()
Dim Ix
For Each Ix In AyNz(SrcMthIx(A, B))
    PushISomSz SrcMthDry, SrcMthIxDr(A, Ix, , C)
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

Sub Z_SrcMthIx()
Dim A$(), Ix%(), O$(), I
A = CurSrc
Ix = SrcMthIx(CurSrc)
For Each I In Ix
    PushI O, A(I)
Next
Brw O
End Sub

Sub Z_IsMthLin()
GoSub Browse
Dim A$
A = "Function IsMthLin(A, Optional B As WhMth) As Boolean"
Ept = True
GoSub Tst
Exit Sub
Tst:
    Act = IsMthLin(A)
    C
    Return
Browse:
Dim L, O$()
For Each L In CurSrc
    If IsMthLin(L) Then
        PushI O, L
    End If
Next
Brw O
End Sub

Function SrcMthIx(A$(), Optional B As WhMth) As Integer()
Dim J%
For J = 0 To UB(A)
    If IsMthLin(A(J), B) Then
        Push SrcMthIx, J
    End If
Next
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
    I = SrcMthIxTopRmkFm(A, I)
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
    Fmix = SrcMthIxTopRmkFm(A, Fm)
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
SrcMthIxLinesWithRmk = ApLines(SrcMthIxTopRmk(A, MthIx), SrcMthIxLines(A, MthIx))
End Function
Function SrcMthIxLines$(A$(), MthIx)
Dim L2%
L2 = SrcMthIxTo(A, MthIx): If L2 = 0 Then Stop
SrcMthIxLines = Join(AyWhFmTo(A, MthIx, L2), vbCrLf)
End Function



Function SrcMthIxTopRmkFm%(A$(), MthIx)
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
SrcMthIxTopRmkFm = M2
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


Function SrcMthLinesWithRmk$(A$(), MthNm)
Dim L%(): L = SrcMthNmIx(A, MthNm)
If Sz(L) = 0 Then Exit Function
Dim MthIx, O$()
For Each MthIx In L
    Push O, SrcMthIxLinesWithRmk(A, MthIx)
Next
SrcMthLinesWithRmk = Join(O, vbCrLf & vbCrLf)
End Function
Function SrcMthIxTopRmk$(A$(), MthIx)
Dim O$(), J%, L$, I%
Dim Fm%: Fm = SrcMthIxTopRmkFm(A, MthIx)

For J = Fm To MthIx - 1
    If IsRmkLin(A(J)) Then Push O, L
Next
SrcMthIxTopRmk = Join(O, vbCrLf)
End Function
Function SrcMthLines$(A$(), MthNm$)
Dim I, O$()
For Each I In AyNz(SrcMthNmIx(A, MthNm))
    Push O, SrcMthIxLines(A, I)
Next
SrcMthLines = Join(O, vbCrLf & vbCrLf)
End Function


Attribute VB_Name = "IdeMthDrPjMdMthNPrm"
Option Explicit
Private Sub ZZ_VbeWsMth5NPrm()
WsVis VbeWsMth5NPrm(CurVbe)
End Sub
Function VbeWsMth5NPrm(A As Vbe) As Worksheet
Set VbeWsMth5NPrm = DrsWs(VbeDrsMth5NPrm(A))
End Function
Function VbeDrsMth5NPrm(A As Vbe) As Drs
Set VbeDrsMth5NPrm = Drs("Pj Md Mdy Ty Mth NPrm", VbeDryMth5NPrm(A))
End Function
Function VbeDryMth5NPrm(A As Vbe) As Variant()
VbeDryMth5NPrm = AyMapFlat(VbePjAy(A), "PjDryMth5NPrm")
End Function
Private Sub ZZ_PjDrymth5NPrm()
Dim O()
O = PjDryMth5NPrm(CurPj)
Stop
End Sub
Private Sub ZZ_SrcDryMth7()
Dim Act()
Act = SrcDryMth7(CurSrc)
Stop
End Sub
Function DrrDry(A As Collection) As Variant()
Dim O(), Dr
For Each Dr In A
    If Not IsArray(Dr) Then Stop
    Push O, Dr
Next
DrrDry = O
End Function

Function MdMth9Dry(A As CodeModule) As Variant()
'Pj Md Mdy Ty Nm Sfx Prm Ret Rmk
'1  2  3   4  5  6   7   8   9
Dim Src$(), Dry()
Src = MdSrc(A)
Dry = SrcDryMth7(Src)
Stop '
End Function
Function DrrCC_CCXDrr(A As Collection, C1, C2) As Collection
Dim O As New Collection
Stop '
Set DrrCC_CCXDrr = O
End Function
Function LinCnt&(Lines)
LinCnt = SubStrCnt(Lines, vbCrLf) + 1
End Function
Function PjDryMth12(A As VBProject) As Variant()
PjDryMth12 = AyMapFlat(PjCdMdAy(A), "MdDryMth12")
End Function
Function PjDryFun12(A As VBProject) As Variant()
PjDryFun12 = AyMapFlat(PjMdAy(A), "MdDryFun12")
End Function

Function PjDrsMth12(A As VBProject) As Drs
PjDrsMth12 = Drs(FnyOf_Mth12, PjDryMth12(A))
End Function
Function VbeDrsMth12(A As Vbe) As Drs
Set VbeDrsMth12 = Drs(FnyOf_Mth12, VbeDryMth12(A))
End Function
Function VbeDrsFun12(A As Vbe) As Drs
Set VbeDrsFun12 = Drs(FnyOf_Mth12, VbeDryFun12(A))
End Function

Private Sub ZZ_VbeWsMth12()
WsVis VbeWsMth12(CurVbe)
End Sub
Function VbeWsMth12(A As Vbe) As Worksheet
Set VbeWsMth12 = DrsWs(VbeDrsMth12(A))
End Function
Function FnyOf_Mth12() As String()
FnyOf_Mth12 = SslSy("Pj Md Mdy Ty Nm Sfx Prm Ret Rmk Lno Cnt Lines")
                    '1  2  3   4  5  6   7   8   9   10  11  12
End Function
Function VbeDryMth12(A As Vbe) As Variant()
VbeDryMth12 = AyMapFlat(VbePjAy(A), "PjDryMth12")
End Function
Function VbeDryFun12(A As Vbe) As Variant()
VbeDryFun12 = AyMapFlat(VbePjAy(A), "PjDryFun12")
End Function
Function MdDryMth12(A As CodeModule) As Variant()
Dim Pj$, Md$, Dry()
Pj = MdPjNm(A)
Md = MdNm(A)
Dry = SrcDryMth10(MdSrc(A))
MdDryMth12 = DryCC_CCXDry(Dry, Pj, Md)
End Function
Function MdDryFun12(A As CodeModule) As Variant()
If Not MdIsStd(A) Then Exit Function
MdDryFun12 = MdDryMth12(A)
End Function
Private Sub ZZ_SrcDryMth10()
Dim Act()
Act = SrcDryMth10(CurSrc)
Stop
End Sub
Function SrcDryMth10(A$()) As Variant()
'Mdy Ty Nm Sfx Prm Ret Rmk Lno Cnt Lines
'1   2  3  4   5   6   7   8   9   10
Dim L, O(), Lin$, M7(), Ay%(), Lines$
Ay = SrcMthIxAy(A)
If Sz(Ay) = 0 Then Exit Function
For Each L In Ay
    Lin = SrcContLin(A, CStr(L))
    M7 = LinMth7DotDr(Lin)
    If Sz(M7) = 0 Then GoTo X
    Lines = SrcMthIxLines(A, L)
    PushAy M7, Array(L + 1, LinCnt(Lines), Lines)
    Push O, M7
X:
Next
SrcDryMth10 = O
End Function
Function SrcDryMth7(A$()) As Variant()
Dim L, O()
For Each L In AyNz(SrcMthLinAy(A))
    PushWithSz O, LinMth7DotDr(L)
Next
SrcDryMth7 = O
End Function
Function LinMth7DotDr(A) As Variant()
'Mdy Ty Nm Sfx Prm Ret Rmk
'1   2  3  4   5   6   7
Dim M$, T$, N$, S$, P$, R$, Rmk$
Dim L$, OptAs$
AyAsg ShiftMdy(A), L
AyAsg ShiftMthShtTy(L), T, L: If T = "" Then Exit Function
AyAsg ShiftNm(L), N, L
AyAsg ShiftTySfxChr(L), S, L
AyAsg ShiftBktStr(L), P, L
AyAsg ShiftX(L, "As"), OptAs, L
If OptAs <> "" Then
    AyAsg ShiftT1(L), R, L: If R = "" Then Stop
End If
If Len(L) > 0 Then
    L = LTrim(L)
    If FstChr(L) <> "'" Then Stop
    Rmk = L
End If
LinMth7DotDr = Array(M, T, N, S, P, R, Rmk)
End Function
Private Sub ZZ_VbeDrymth5NPrm()
Dim O()
O = VbeDryMth5NPrm(CurVbe)
Stop
End Sub
Function PjDryMth5NPrm(A As VBProject) As Variant()
PjDryMth5NPrm = AyMapFlat(PjCdMdAy(A), "MdDryMth5NPrm")
End Function

Function AyMapFlat(Ay, ItmAyFun$)
AyMapFlat = AyMapFlatInto(Ay, ItmAyFun, EmpAy)
End Function
Function AyMapFlatInto(Ay, ItmAyFun$, OIntoAy)
Dim O, J&, M
O = OIntoAy: Erase O
For J = 0 To UB(Ay)
    M = Run(ItmAyFun, Ay(J))
    PushAy O, M
Next
AyMapFlatInto = O
End Function
Private Sub ZZ_MdDryMth5NPrm()
Dim O()
O = MdDryMth5NPrm(CurMd)
Stop
End Sub
Function MdDryMth5NPrm(A As CodeModule) As Variant()
Dim Dry(), Pj$, Md$
Dry = SrcDryMth3NPrm(MdSrc(A))
Pj = MdPjNm(A)
Md = MdNm(A)
MdDryMth5NPrm = DryCC_CCXDry(Dry, Pj, Md)
End Function

Private Sub ZZ_SrcDryMth3NPrm()
Dim O()
O = SrcDryMth3NPrm(CurSrc)
Stop
End Sub
Function SrcDryMth3NPrm(A$()) As Variant()
Dim O(), L
For Each L In A
    PushWithSz O, LinMth3DotNPrmDr(L)
Next
SrcDryMth3NPrm = O
End Function
Function LinMth3DotNPrmDr(A) As Variant()
Dim M$, T$, N$, C$, P$, NP%, L$
Dim Brk$()
AyAsg ShiftMthBrk(A), Brk, L
If Brk(2) = "" Then Exit Function
AyAsg Brk, M, T, N
AyAsg ShiftTySfxChr(L), C, L
AyAsg ShiftBktStr(L), P$
If Len(P) > 0 Then
    NP = SubStrCnt(P, ",") + 1
End If
LinMth3DotNPrmDr = Array(M, T, N, NP)
End Function
Sub PushWithSz(O, Ay)
If Not IsArray(Ay) Then Stop
If Sz(Ay) = 0 Then Exit Sub
Push O, Ay
End Sub
Function DryCC_CCXDry(A, C1, C2) As Variant()
Dim Dr, O(), CCAy()
If Sz(A) = 0 Then Exit Function
CCAy = Array(C1, C2)
For Each Dr In A
    Push O, AyInsAy(Dr, CCAy)
Next
DryCC_CCXDry = O
End Function
Private Sub Z_AyInsAy()
Dim Act, Exp, A(), B(), At&
A = Array(1, 2, 3, 4)
B = Array("X", "Z")
At = 1
Exp = Array(1, "X", "Z", 2, 3, 4)

Act = AyInsAy(A, B, At)
Ass AyIsEq(Act, Exp)
End Sub
Function AyIsEq(A, B) As Boolean
If Not IsArray(A) Then Stop
If Not IsArray(B) Then Stop
Dim U&, J&
U = UB(A)
If U <> UB(B) Then Exit Function
For J = 0 To U
    If A(J) <> B(J) Then Exit Function
Next
AyIsEq = True
End Function
Private Sub ZZ_AyReSzAt()
Dim Ay(), At&, Cnt&, Act, Exp
Ay = Array(1, 2, 3)
At = 1
Cnt = 3
Exp = Array(1, Empty, Empty, Empty, 2, 3)
Act = AyReSzAt(Ay, At, Cnt)
Ass AyIsEq(Act, Exp)
End Sub
Function AyReSzAt(A, At&, Optional Cnt& = 1)
If Cnt < 1 Then Stop
Dim O, U&, J&, F&, T&
O = A
U = UB(A)
'----------
Dim NU&, FOS&, TOS&, UM&
NU = U + Cnt
FOS = U
TOS = NU
UM = U - At
ReDim Preserve O(NU)
For J = 0 To UM
    F = FOS - J
    T = TOS - J
    Asg O(F), O(T)
    O(F) = Empty
Next
AyReSzAt = O
End Function
Function AyInsAy(A, B, Optional At&)
Dim O, NA&, NB&, J&
NA = Sz(A)
NB = Sz(B)
O = AyReSzAt(A, At, NB)
For J = 0 To NB - 1
    Asg B(J), O(At + J)
Next
AyInsAy = O
End Function

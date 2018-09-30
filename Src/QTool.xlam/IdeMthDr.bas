Attribute VB_Name = "IdeMthDr"
Option Explicit
Private Sub ZZ_SrcMth7Dry()
Dim Act()
Act = SrcMth7Dry(CurSrc)
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

Function DrrCC_CCXDrr(A As Collection, C1, C2) As Collection
Dim O As New Collection
Stop '
Set DrrCC_CCXDrr = O
End Function
Function PjMth12Dry(A As VBProject) As Variant()
PjMth12Dry = AyMapFlat(PjMdAy(A), "MdMth12Dry")
End Function
Function PjFun12Dry(A As VBProject) As Variant()
PjFun12Dry = AyMapFlat(PjMdAy(A), "MdDryFun12")
End Function

Function PjMth12Drs(A As VBProject) As Drs
PjMth12Drs = Drs(Mth12DrFny, PjMth12Dry(A))
End Function
Function CurVbeSrc() As String()
CurVbeSrc = VbeSrc(CurVbe)
End Function
Function PjSrc(A As VBProject) As String()
Dim C As VBComponent
For Each C In A.VBComponents
    PushAy PjSrc, MdSrc(C.CodeModule)
Next
End Function
Sub ZZ_VbeMth12Ws()
WsVis VbeMth12Ws(CurVbe)
End Sub
Sub ZZ_SrcMth10Dry()
Dim Act()
Act = SrcMth10Dry(CurSrc)
Stop
End Sub

Sub Z_LinMth7Dr()
Dim A$
A = "Sub MdAppLy(A As CodeModule, Ly$())"
Ept = ApSy("", "Sub", "MdAppLy", "", "A As CodeModule, Ly$()", "", "")
GoSub Tst
Exit Sub
Tst:
    Act = LinMth7Dr(A)
    C
    Return
End Sub

Function LinMth7Dr(A) As String()
'Mdy Ty Nm Sfx Prm Ret ShtRmk
'1   2  3  4   5   6   7
'Dim M$, T$, N$, S$, P$, R$, Rmk$
'Dim L$, OptAs$
'AyAsg ShfMdy(A), M, L
'AyAsg ShfMthShtTy(L), T, L: If T = "" Then Exit Function
'AyAsg ShfNm(L), N, L
'AyAsg ShfTySfxChr(L), S, L
'AyAsg ShfBktStr(L), P, L
'AyAsg ShfX(L, "As"), OptAs, L
'If OptAs <> "" Then
'    AyAsg ShfTerm(L), R, L: If R = "" Then Stop
'End If
'If Len(L) > 0 Then
'    L = LTrim(L)
'    If FstChr(L) <> "'" Then Stop
'    Rmk = L
'End If
'LinMth7Dr = ApSy(M, T, N, S, P, R, Rmk)
End Function



Sub PushWithSz(O, Ay)
If Not IsArray(Ay) Then Stop
If Sz(Ay) = 0 Then Exit Sub
Push O, Ay
End Sub
Private Sub Z_AyInsAy()
Dim Act, Exp, A(), B(), At&
A = Array(1, 2, 3, 4)
B = Array("X", "Z")
At = 1
Exp = Array(1, "X", "Z", 2, 3, 4)

Act = AyInsAy(A, B, At)
Ass IsEqAy(Act, Exp)
End Sub
Private Sub ZZ_AyReSzAt()
Dim Ay(), At&, Cnt&, Act, Exp
Ay = Array(1, 2, 3)
At = 1
Cnt = 3
Exp = Array(1, Empty, Empty, Empty, 2, 3)
Act = AyReSzAt(Ay, At, Cnt)
Ass IsEqAy(Act, Exp)
End Sub

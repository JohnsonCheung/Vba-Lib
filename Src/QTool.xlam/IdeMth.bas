Attribute VB_Name = "IdeMth"
Option Explicit
Sub MthEnsPrv(A As Mth)
Dim F%(): F = MthLnoAy(A)
Dim F1%(), N$
Dim L2$, J%, L1$, L$
N = MthDNm(A)
For J = 0 To UB(F)
    L = F(J)
    L1 = A.Md.Lines(L, 1)
    L2 = MthLin_EnsPrv(L1)
    If L1 <> L2 Then
        Debug.Print FmtQQ("MthEnsPub: Md(?) Lin(?) RplBy(?)", N, L, L2)
        A.Md.ReplaceLine L, L2
    End If
Next
End Sub
Sub MthEnsPub(A As Mth)
Dim F%(): F = MthLnoAy(A)
Dim F1%(), N$
Dim L2$, J%, L1$, L$
N = MthDNm(A)
For J = 0 To UB(F)
    L = F(J)
    L1 = A.Md.Lines(L, 1)
    L2 = MthLin_EnsPub(L1)
    If L1 <> L2 Then
        Debug.Print FmtQQ("MthEnsPub: Md(?) Lin(?) RplBy(?)", N, L, L2)
        A.Md.ReplaceLine L, L2
    End If
Next
End Sub
Function MthLin_EnsPrv$(A)
MthLin_EnsPrv = "Private " & RmvMdy(A)
End Function
Function MthLin_EnsPub$(A)
MthLin_EnsPub = RmvMdy(A)
End Function
Private Sub Z_MthEnsPub()
Dim M As Mth: Set M = Mth(Md("ZZModule"), "YYA")
MthEnsPrv M: Ass MthLin(M) = "Private Property Get ZZA()"
MthEnsPub M:  Ass MthLin(M) = "Property Get ZZA()"
End Sub
Sub MthRmk(A As Mth)
Dim P() As FTNo: P = MthCxtFT(A)
Dim J%
For J = UB(P) To 0 Step -1
    MthCxtFT_Rmk A, P(J)
Next
End Sub
Sub MthUnRmk(A As Mth)
Dim P() As FTNo: P = MthCxtFT(A)
Dim J%
For J = UB(P) To 0 Step -1
    MthCxtFT_UnRmk A, P(J)
Next
End Sub
Private Sub ZZ_MthRmk()
Dim M As Mth: Set M = Mth(Md("ZZModule"), "YYA")
            Ass LinesVbl(MthLines(M)) = "Property Get ZZA()|End Property||Property Let YYA(V)||End Property"
MthRmk M:   Ass LinesVbl(MthLines(M)) = "Property Get ZZA()|Stop '|End Property||Property Let YYA(V)|Stop '|'|End Property"
MthUnRmk M: Ass LinesVbl(MthLines(M)) = "Property Get ZZA()|End Property||Property Let YYA(V)||End Property"
End Sub
Function MthBrkOptFny(A As MthBrkOpt) As String()
MthBrkOptFny = SslSy("MthIx Cnt Lin Pj Md")
End Function
Function MthKeyDrFny() As String()
MthKeyDrFny = SslSy("PjNm MdNm Priority Nm Ty Mdy")
End Function
Function MthDryWhDup(MthDry(), Optional IsSamMthBdyOnly As Boolean) As String()
'MthFNy is in format of Pj Md ShtMdy ShtTy Nm Lines
If Sz(MthDry) = 0 Then Exit Function
Stop '
Dim Ny$(): Ny = DryStrCol(MthDry, 4)
Dim A1$(): A1 = AySrt(Ny)
Dim O$(), M$(), J&, Nm$
Dim L$ ' LasFunNm
L = MthFNm_Nm(A1(0))
Push M, A1(0)
For J = 1 To UB(A1)
    Nm = MthFNm_Nm(A1(J))
    If L = Nm Then
        Push M, A1(J)
    Else
        L = Nm
        If Sz(M) = 1 Then
            M(0) = A1(J)
        Else
            PushAy O, M
            Erase M
        End If
    End If
Next
If Sz(M) > 1 Then
    PushAy O, M
End If
MthDryWhDup = O
End Function
Private Sub Z_MthPfx()
Ass MthPfx("Add_Cls") = "Add"
End Sub
Private Sub ZZ_MthPfx()
Dim Ay$(): Ay = VbeMthNy(CurVbe)
Dim Ay1$(): Ay1 = AyMapSy(Ay, "MthPfx")
WsVis AyabWs(Ay, Ay1)
End Sub
Function MthPfx$(MthNm$)
Dim A0$
    A0 = Brk1(RmvPfxAy(MthNm, SplitVBar("ZZ_|Z_")), "__").S1
With Brk2(A0, "_")
    If .S1 <> "" Then
        MthPfx = .S1
        Exit Function
    End If
End With
Dim P2%
Dim Fnd As Boolean
    Dim C%
    Fnd = False
    For P2 = 2 To Len(A0)
        C = Asc(Mid(A0, P2, 1))
        If AscIsLCase(C) Then Fnd = True: Exit For
    Next
'---
    If Not Fnd Then Exit Function
Dim P3%
Fnd = False
    For P3 = P2 + 1 To Len(A0)
        C = Asc(Mid(A0, P3, 1))
        If AscIsUCase(C) Or AscIsDigit(C) Then Fnd = True: Exit For
    Next
'--
If Fnd Then
    MthPfx = Left(A0, P3 - 1)
    Exit Function
End If
MthPfx = MthNm
End Function
Function MthKd$(MthTy$)
Select Case MthTy
Case "Function": MthKd = "Fun"
Case "Sub": MthKd = "Sub"
Case "Property Get", "Property Get", "Property Let": MthKd = "Prp"
End Select
End Function
Function MthDNm$(A As Mth)
MthDNm = MdDNm(A.Md) & "." & A.Nm
End Function
Function MthDNmLines$(A)
MthDNmLines = MthLines(DMth(A))
End Function
Function MthFul$(MthNm)
MthFul = VbeMthMdDNm(CurVbe, MthNm)
End Function
Function MthFuly(MthNm) As String()
MthFuly = VbeMthMdDNy(CurVbe, MthNm)
End Function
Function MthNmMdDNy(A) As String()
MthNmMdDNy = CurVbeMthMdDNy(A)
End Function
Function MthNmMd(A) As CodeModule '
Dim O As CodeModule
Set O = CurMd
If MdHasMth(O, A) Then Set MthNmMd = O: Exit Function
Dim N$
N = MthFul(A)
If N = "" Then
    Debug.Print FmtQQ("Mth[?] not found in any Pj")
    Exit Function
End If
Set MthNmMd = Md(N)
End Function
Function MthDNm_Nm$(A)
Dim Ay$(): Ay = Split(A, ".")
Dim Nm$
Select Case Sz(Ay)
Case 1: Nm = Ay(0)
Case 2: Nm = Ay(1)
Case 3: Nm = Ay(2)
Case Else: Stop
End Select
MthDNm_Nm = Nm
End Function
Function MthFNm$(A As Mth)
MthFNm = A.Nm & ":" & MdDNm(A.Md)
End Function
Function MthFNm_Mth(A) As Mth
Set MthFNm_Mth = DMth(MthFNm_MthDNm(A))
End Function
Function MthFNm_MthDNm$(A)
With Brk(A, ":")
    MthFNm_MthDNm = .S2 & "." & .S1
End With
End Function
Function MthFNm_Nm$(A$)
MthFNm_Nm = Brk(A, ":").S1
End Function
Function MthLno(A As Mth) As Integer()
MthLno = MdMthLno(A.Md, A.Nm)
End Function
Function MthLnoAy(A As Mth) As Integer()
MthLnoAy = AyAdd1(SrcMthNmIx(MdSrc(A.Md), A.Nm))
End Function
Function MthRmkFC(A As Mth) As FmCnt()
MthRmkFC = SrcMthRmkFC(MdSrc(A.Md), A.Nm)
End Function
Function MthKy_Sq(A$()) As Variant()
Dim O(), J%
ReDim O(1 To Sz(A) + 1, 1 To 6)
SqSetRow O, 1, MthKeyDrFny
For J = 0 To UB(A)
    SqSetRow O, J + 2, Split(A(J), ":")
Next
MthKy_Sq = O
End Function
Function MthMayLCC(A As Mth) As MayLCC
Dim L%, C As MayLCC
Dim M As CodeModule
Set M = A.Md
For L = M.CountOfDeclarationLines + 1 To M.CountOfLines
    Set C = LinMayLCC(M.Lines(L, 1), A.Nm, L)
    If C.Som Then
        Set MthMayLCC = SomLCC(C.LCC)
        Exit Function
    End If
Next
End Function
Function MthANy_MthNmItr(A) As Collection
Dim O As New Collection, J&
For J = 0 To UB(A)
    ItrPushNoDup O, A(J)
Next
Set MthANy_MthNmItr = O
End Function
Function MthLin$(A As Mth)
MthLin = SrcMthLin(MdBdyLy(A.Md), A.Nm)
End Function
Function MthANm_MthNm$(A)
MthANm_MthNm = TakBefOrAll(A, ":")
End Function
Function MthBNm_MthNm$(A)
MthBNm_MthNm = MthANm_MthNm(MthBNm_MthANm(A))
End Function
Function MthBNm_MdNm$(A)
MthBNm_MdNm = TakBefMust(A, ".")
End Function
Function MthBNm_MthANm$(A)
MthBNm_MthANm = TakAftMust(A, ".")
End Function
Function MthLinCnt%(A As Mth)
MthLinCnt = FmCntAyLinCnt(MthFC(A))
End Function
Function MthMdNm$(A As Mth)
MthMdNm = MdNm(A.Md)
End Function
Function MthMdDNm$(A As Mth)
MthMdDNm = MdDNm(A.Md)
End Function
Function MthNm$(A As Mth)
MthNm = A.Nm
End Function
Function MthNm_CmpLy(A, Optional InclSam As Boolean) As String()
Dim N$(): N = MthNm_DupMthFNy(A)
If Sz(N) > 1 Then
    MthNm_CmpLy = DupMthFNyGp_CmpLy(N, InclSam:=InclSam)
End If
End Function
Function MthNm_DupMthFNy(A) As String()
Stop '
'MthNm_DupMthFNy = VbeFunFNm(CurVbe, FunPatn:="^" & A & "$")
End Function
Function MthPjNm$(A As Mth)
MthPjNm = MdPjNm(A.Md)
End Function
Function MthTy_IsVdt(A) As Boolean
MthTy_IsVdt = AyHas(MthTyAy, A)
End Function
Function MthShtKd$(MthKd)
Dim O$
Select Case MthKd
Case "Sub": O = MthKd
Case "Function": O = "Fun"
Case "Property": O = "Prp"
End Select
MthShtKd = O
End Function
Function MthFC(A As Mth) As FmCnt()
MthFC = SrcMthNmFC(MdBdyLy(A.Md), A.Nm)
End Function
Function MthEndLin$(MthLin$)
Dim A$
A = LinMthKd(MthLin): If A = "" Then Stop
MthEndLin = "End " & A
End Function
Function MthTyAy() As String()
Static O$(4), A As Boolean
If Not A Then
    A = True
    O(0) = "Property Get"
    O(1) = "Property Let"
    O(2) = "Property Set"
    O(3) = "Sub"
    O(4) = "Function"
End If
MthTyAy = O
End Function
Function MthKdAy() As String()
Static O$(2), A As Boolean
If Not A Then
    A = True
    O(1) = "Sub"
    O(0) = "Function"
    O(2) = "Property"
End If
MthKdAy = O
End Function
Function MthShtTyAy() As String()
Static O$(4), A As Boolean
If Not A Then
    A = True
    O(0) = "Get"
    O(1) = "Let"
    O(2) = "Set"
    O(3) = "Sub"
    O(4) = "Fun"
End If
MthShtTyAy = O
End Function
Function MthNy() As String()
MthNy = CurVbeMthNy
End Function
Function MthNyWh(A As WhPjMth) As String()
MthNyWh = VbeMthNyWh(CurVbe, A)
End Function
Sub MthBrkAsg(A As Mth, OMdy$, OMthTy$)
Dim L$: L = MthLin(A)
OMdy = TakMdy(L)
OMthTy = LinMthTy(L)
End Sub
Sub MthGo(A As Mth)
MdGoMayLCC A.Md, MthMayLCC(A)
End Sub
Function MthCpyPrm_Cpy(A As MthCpyPrm)
MthCpy A.SrcMth, A.ToMd
End Function
Sub MthRmv(A As Mth, Optional IsSilent As Boolean)
Dim X() As FmCnt: X = MthRmkFC(A)
MdRmvFC A.Md, X
If Not IsSilent Then
    Debug.Print FmtQQ("MthRmv: Mth(?) of LinCnt(?) is deleted", MthDNm(A), FmCntAyLinCnt(X))
End If
End Sub
Sub MthRpl(A As Mth, By$)
MthRmv A
MdAppLines A.Md, By
End Sub
Private Sub Z_MthFC()
Dim M As Mth: Set M = Mth(Md("ZZModule"), "YYA")
Dim Act() As FmCnt: Act = MthFC(M)
Ass Sz(Act) = 2
Ass Act(0).FmLno = 5
Ass Act(0).Cnt = 7
Ass Act(1).FmLno = 13
Ass Act(1).Cnt = 15
End Sub
Private Sub Z_MthRmv()
Const N$ = "ZZModule"
Dim M As CodeModule
Dim M1 As Mth, M2 As Mth
GoSub Crt
Set M = Md(N)
Set M1 = Mth(M, "ZZRmv1")
Set M2 = Mth(M, "ZZRmv2")
MthRmv M1
MthRmv M2
MdEndTrim M
If M.CountOfLines <> 1 Then MsgBox M.CountOfLines
MdDlt M
Exit Sub
Crt:
    CurPjDltMd N
    Set M = CurPjEnsMd(N)
    MdAppLines M, RplVBar("Property Get ZZRmv1()||End Property||Function ZZRmv2()|End Function||'|Property Let ZZRmv1(V)|End Property")
    Return
End Sub
Sub MthDicB_AssKeysIsBNm(A As Dictionary)
Dim K
For Each K In A.Keys
    If InStr(K, ".") = 0 Then Stop
Next
End Sub
Function Mth12DrFny() As String()
Mth12DrFny = SslSy("Pj Md Mdy Ty Nm Sfx Prm Ret Rmk Lno Cnt Lines")
                    '1  2  3   4  5  6   7   8   9   10  11  12
End Function
Function MthDotNTM$(MthDot$)
'MthDot is a string with last 3 seg as Mdy.ShtTy.Nm
'MthNTM is a string with last 3 seg as Nm:ShtTy.Mdy
Dim Ay$(), Nm$, ShtTy$, Mdy$
Ay = SplitDot(MthDot)
AyAsg AyPop(Ay), Ay, Nm
AyAsg AyPop(Ay), Ay, ShtTy
AyAsg AyPop(Ay), Ay, Mdy
Push Ay, FmtQQ("?:?.?", Nm, ShtTy, Mdy)
MthDotNTM = JnDot(Ay)
End Function
Function MthShtTy$(MthTy)
Dim O$
Select Case MthTy
Case "Sub": O = MthTy
Case "Function": O = "Fun"
Case "Property Get": O = "Get"
Case "Property Let": O = "Let"
Case "Property Set": O = "Set"
End Select
MthShtTy = O
End Function
Function MthBrkDot$(MthBrk$())
If MthBrk(2) = "" Then Exit Function
MthBrkDot = JnDot(MthBrk)
End Function
Function MthCpy(A As Mth, ToMd As CodeModule, Optional IsSilent As Boolean) As Boolean
If MdHasMth(ToMd, A.Nm) Then
    Debug.Print FmtQQ("MthCpy_ToMd: Fm-Mth(?) is Found in To-Md(?)", A.Nm, MdNm(ToMd))
    MthCpy = True
    Exit Function
End If
If ObjPtr(A.Md) = ObjPtr(ToMd) Then
    Debug.Print FmtQQ("MthCpy: Fm-Mth-Md(?) cannot be To-Md(?)", MthMdNm(A), MdNm(ToMd))
    MthCpy = True
    Exit Function
End If
MdAppLines ToMd, MthLines(A)
If Not IsSilent Then
    Debug.Print FmtQQ("MthCpy: Mth(?) is copied ToMd(?)", MthDNm(A), MdDNm(ToMd))
End If
End Function
Sub MthAyMov(A() As Mth, ToMd As CodeModule)
AyDoXP A, "MthMov", ToMd
End Sub
Sub MthMov(A As Mth, ToMd As CodeModule)
If MthCpy(A, ToMd, IsSilent:=True) Then Exit Sub
MthRmv A, IsSilent:=True
Debug.Print FmtQQ("MthMov: Mth(?) is moved to Md(?)", MthDNm(A), MdDNm(ToMd))
End Sub
Function MthLines$(A As Mth)
MthLines = SrcMthLines(MdBdyLy(A.Md), A.Nm)
End Function
Function MthLinesWithRmk$(A As Mth)
MthLinesWithRmk = SrcMthLinesWithRmk(MdBdyLy(A.Md), A.Nm)
End Function

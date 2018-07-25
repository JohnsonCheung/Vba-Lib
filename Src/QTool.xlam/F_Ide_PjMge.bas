Attribute VB_Name = "F_Ide_PjMge"
Option Explicit
Const LblOf_FmPj$ = "Mge from pj"
Const LblOf_ToPj$ = "Mge into pj"

Sub AA()
ZZ_DoMgePj
End Sub
Sub AAAA()
ZZ_DicAB_DryMIS
End Sub
Sub AAA()
ZZ_FTPjMSq
End Sub
Function BNmDIF_Dr(BNmDIF$, DicA As Dictionary, DicB As Dictionary) As Variant()
Dim Dr(), Ay$(), ToMthLinesAy(), ANm$
Dr = BNmMIS_Dr(BNmDIF, DicA)
Ay = Split(BNmDIF, ".")
ANm = Ay(1)
ToMthLinesAy = DicB(ANm)
BNmDIF_Dr = AyAddAp(Dr, ToMthLinesAy)
End Function

Function BNmMIS_Dr(BNmMIS, DicA As Dictionary) As Variant()
Dim Ay$(), FmMdNm$, MthNm$
Dim ToMdNm$, MthShtTy$, ShtMdy$, FmMthLines$
Ay = Split(BNmMIS, "."): If Sz(Ay) <> 2 Then Stop
MthNm = TakBef(Ay(1), ":")
FmMthLines = DicA(BNmMIS)
BNmMIS_Dr = Array(FmMdNm, ToMdNm, MthNm, MthShtTy, ShtMdy, FmMthLines)
End Function

Private Function CellOf_FmPj(A As Worksheet) As Range
Set CellOf_FmPj = WsRC(A, 3, 1)
End Function

Private Function CellOf_FmPjFormula(A As Worksheet) As Range
Set CellOf_FmPjFormula = WsRC(A, 3, 5)
End Function

Private Function CellOf_FmPjLbl(A As Worksheet) As Range
Set CellOf_FmPjLbl = WsRC(A, 2, 1)
End Function

Private Function CellOf_ToPj(A As Worksheet) As Range
Set CellOf_ToPj = WsRC(A, 3, 2)
End Function

Private Function CellOf_ToPjFormula(A As Worksheet) As Range
Set CellOf_ToPjFormula = WsRC(A, 3, 6)
End Function

Private Function CellOf_ToPjLbl(A As Worksheet) As Range
Set CellOf_ToPjLbl = WsRC(A, 2, 2)
End Function

Function DicMaxValSz%(A As Dictionary)
'MthDic is DicOf_MthNm_zz_MthLinesAy
'MaxMthCnt is max-of-#-of-method per MthNm
Dim O%, K
For Each K In A.Keys
    O = Max(O, Sz(A(K)))
Next
DicMaxValSz = O
End Function

Sub DoMgePj()
Dim W As Worksheet: Set W = WsOf_PjMge
PjMgeWs_Bld W
MthCpyAy_Cpy PjMgeWs_MthCpyAy(W)
End Sub

Sub Doc()
'#BNmMIS is Method-B-Nm-of-Missing.
'           Missing means the method is found in FmPj, but not ToPj
'#DicA is MthDic-of-MthANm-zz-MthLinesAy. It comes from ToPj (#ToA)
'#DicB is MthDic-of-MthBNm-zz-MthLines.   It comes from FmPj (#FmB)
'#DicAB is DicA and DicB
'#ANm is method-a-name, NNN or NNN:YYY
'        If the method is Sub|Fun, just MthNm
'        If the method is Prp    ,      MthNm:MthShtTy
'        It is from ToPj (#ToA)
'        One MthANm will have one or more MthLines
'#BNm is method-b-name, MMM.NNN or MMM.NNN:YYY
'        MdNm.MthNm[:MthShtTy]
'        It is from FmPj (#BFm)
'        One MthBNm will have only one MthLines
'#ATo is used to remember: method-A is coming from ToPj
'#BFm is used to remember: method-A is coming from FmPj
End Sub

Function DicAB_BNyDIF(DicA As Dictionary, DicB As Dictionary) As String()

End Function

Function DicAB_DryDIF(DicA As Dictionary, DicB As Dictionary) As Variant()
Dim O(), U%, MthBNyOfDif$(), J%
MthBNyOfDif = DicAB_BNyDIF(DicA, DicB)
U = UB(MthBNyOfDif): If U < 0 Then Exit Function
ReDim O(U)
For J = 0 To U
    O(J) = BNmDIF_Dr(MthBNyOfDif(J), DicA, DicB)
Next
DicAB_DryDIF = O
End Function

Sub ZZ_DicAB_DryMIS()
Dim A(): A = DicAB_DryMIS(ZZDicA, ZZDicB)
Stop
End Sub
Function DicAB_DryMIS(DicA As Dictionary, DicB As Dictionary) As Variant()
Dim O(), BNmsMIS As Collection, J%, BNm
Set BNmsMIS = DicAB_BNmsMIS(DicA, DicB)
ReDim O(BNmsMIS.Count)
J = 0
For Each BNm In BNmsMIS
    O(J) = BNmMIS_Dr(BNm, DicA)
    J = J + 1
Next
DicAB_DryMIS = O
End Function

Function DicAB_BNmsMIS(DicA As Dictionary, DicB As Dictionary) As Collection
Dim BNm, O As New Collection
For Each BNm In DicA.Keys
Next
Set DicAB_BNmsMIS = O
End Function

Private Function CNmProperMdNm$(A$)
'Given a [Mth}, return the MdNm which the Mth should be copied to
Stop '
End Function

Private Function FTPjMSq(FmPj As VBProject, ToPj As VBProject) As Variant()
Dim D As Dictionary
Dim MaxToMthCnt% ' ToPj is using MthNm as key to get all mth (no matter which md and mdy)
                 ' MaxToMthCnt is MaxMthCnt-of-ToPj
Dim DicA As Dictionary, DicB As Dictionary
Dim Fny$(), A(), B(), Dry()
Set D = PjDicA(ToPj)
MaxToMthCnt = DicMaxValSz(D)
Set DicA = PjDicA(FmPj)
Set DicB = PjDicB(ToPj)
Fny = FnyOf_PjMge(MaxToMthCnt)
A = DicAB_DryMIS(DicA, DicB)
B = DicAB_DryDIF(DicA, DicB)
Stop
Dry = AyAddAp(Array(Fny), A, B)
FTPjMSq = DrySq(Dry)
End Function

Function FnyOf_PjMge(MaxMthCnt%) As String()
Const C$ = "FmMd ToMd Mth Sel FmMth "
Dim O$
If MaxMthCnt = 1 Then
    O = C & "ToMth"
Else
    Dim Ay$(), J%
    ReDim Ay$(MaxMthCnt - 1)
    For J = 0 To MaxMthCnt - 1
        Ay(J) = "ToMth" & (J + 1)
    Next
    O = C & JnSpc(Ay)
End If
FnyOf_PjMge = SslSy(O)
End Function

Function JnSpc$(A)
JnSpc = Join(A, " ")
End Function

Function MdMthLyItr(A As CodeModule) As Collection
Set MdMthLyItr = SrcMthLyItr(MdSrc(A))
End Function
Function MdDicA(A As CodeModule) As Dictionary
Set MdDicA = SrcDicA(MdSrc(A))
End Function
Function MdDicB(A As CodeModule) As Dictionary
Set MdDicB = DicAddKeyPfx(MdDicA(A), MdNm(A) & ".")
End Function

Function MthCpy(A As Mth, ToMd As CodeModule) As MthCpy
Dim O As New MthCpy
Set MthCpy = O.Init(A, ToMd)
End Function

Function MthCpyAy_Cpy(A() As MthCpy)
If Sz(A) = 0 Then Exit Function
Dim M As MthCpy, I
For Each I In A
    Set M = I
    MthCpy M.SrcMth, M.ToMd
Next
End Function

Function PjMgeSq_MthCpy(Sq, R&, FmPj As VBProject, ToPj As VBProject) As MthCpy()
Dim MthNm$, FmMdNm$, ToMdNm$
Dim SrcMth As Mth, ToMd As CodeModule, M As MthCpy
Dim Mth As Mth
    FmMdNm = Sq(R, 1)
    ToMdNm = Sq(R, 2)
    MthNm = Sq(R, 3)  '<=======
    Set SrcMth = Mth(PjMd(FmPj, FmMdNm), MthNm)
    Set M.SrcMth = Mth
    Set M.ToMd = PjMd(ToPj, ToMdNm)
PjMgeSq_MthCpy = MthCpy(M, ToMd)
End Function

Function PjMgeSq_MthCpyAy(Sq, FmPj As VBProject, ToPj As VBProject) As MthCpy()
Dim R&, IsSel, O() As MthCpy
Dim MthNm$, FmMd As CodeModule
For R = 1 To UBound(Sq, 1)
    IsSel = Sq(R, 4)
    If IsSel = "X" Then
        PushObj O, PjMgeSq_MthCpy(Sq, R, FmPj, ToPj)
    End If
Next
PjMgeSq_MthCpyAy = O
End Function

Sub PjMgeWs_Bld(A As Worksheet)
If Not WsIsPjMgeWs(A) Then
    MsgBox FmtQQ("Given: Ws(?) is not PjMgeWs, which has A1=[Mge from Pj] and A2=[Mge Into Pj]", A.Name)
    Exit Sub
End If
'-- Set Bdr
    RgBdrAround RgOf_FmPj(A)
    RgBdrAround RgOf_ToPj(A)
'-- Set Formula
    'C1 = FmPj.Md  C2 = ToPj.Md | C3 = MthNm | C4 = Sel | C5 = FmMthLines | C6 = ToMthLines (Only for DifMth, not MisMth)
    WsRC(A, 3, 6).Formula = "=$A$3"
    WsRC(A, 3, 7).Formula = "=$B$3"
'--- Protect
    PjMgeWs_SetProtect A
Dim ToPj As VBProject: Set ToPj = PjMgeWs_ToPj(A)
Dim FmPj As VBProject: Set FmPj = PjMgeWs_FmPj(A)
PjMgeWs_SetPjErMsg A, FmPj, ToPj
If IsNothing(ToPj) Then Exit Sub
If IsNothing(FmPj) Then Exit Sub
CellPutSq WsRC(A, 3, 1), FTPjMSq(FmPj, ToPj)
End Sub

Private Function PjMgeWs_FmPj(A As Worksheet) As VBProject
Dim P$: P = CellOf_FmPj(A).Value
If Not VbeHasPj(CurVbe, P) Then Exit Function
Set PjMgeWs_FmPj = Pj(P)
End Function

Function PjMgeWs_MthCpyAy(A As Worksheet) As MthCpy()
Dim Lo As ListObject
If A.ListObjects.Count = 0 Then Exit Function
Set Lo = A.ListObjects(1)
Dim ToPj As VBProject
Dim FmPj As VBProject
    Set ToPj = PjMgeWs_ToPj(A)
    Set FmPj = PjMgeWs_ToPj(A)
SelRg_SetXorEmpty Lo.ListColumns("Sel").DataBodyRange
PjMgeWs_MthCpyAy = PjMgeSq_MthCpyAy(Lo.DataBodyRange.Value, FmPj, ToPj)
End Function

Private Sub PjMgeWs_SetPjErMsg(A As Worksheet, FmPj As VBProject, ToPj As VBProject)
PjMgeWs_SetPjErMsg__X A, FmPj, IsFmPj:=True
PjMgeWs_SetPjErMsg__X A, ToPj, IsFmPj:=False
End Sub

Private Sub PjMgeWs_SetPjErMsg__X(A As Worksheet, Pj As VBProject, IsFmPj As Boolean)
Dim C%, Colr%
If IsFmPj Then
    C = CellOf_FmPj(A).Column
    Colr = 1
Else
    C = CellOf_ToPj(A).Column
    Colr = 2
End If
Dim R As Range
Dim At As Range
    Set At = WsRC(A, 1, C)
    Set R = WsRCRC(A, 1, C, 3, C)
If IsNothing(Pj) Then
    At.Value = "Project Not Found"
    R.Interior.Color = Colr
Else
    R.Interior.Color = 0
    At.Clear
End If
End Sub

Sub PjMgeWs_SetProtect(A As Worksheet)

End Sub

Private Function PjMgeWs_ToPj(A As Worksheet) As VBProject
Dim P$: P = CellOf_ToPj(A).Value
If Not VbeHasPj(CurVbe, P) Then Exit Function
Set PjMgeWs_ToPj = Pj(P)
End Function
Function ItrMap(A, MapFunNm$) As Collection
Dim O As New Collection, I
For Each I In A
    O.Add Run(MapFunNm, I)
Next
Set ItrMap = O
End Function
Function PjMthLyItr(A As VBProject) As Collection
Dim IItr As Collection: Set IItr = ItrMap(PjMdAy(A), "MdMthLyItr")
Set PjMthLyItr = IItrItr(IItr)
End Function

Function PjDicA(A As VBProject) As Dictionary
Dim LyItr As Collection, MthANm$, M()
Dim Ly$(), I, O As New Dictionary
Set LyItr = PjMthLyItr(A)
For Each I In LyItr
    Ly = I
    MthANm = MthLin_MthANm(SrcContLin(Ly, 0))
    If MthANm = "" Then Stop
    If O.Exists(MthANm) Then
        M = O(MthANm)
        Push M, Ly
        O(MthANm) = M
    Else
        O.Add MthANm, Array(Ly)
    End If
Next
Set PjDicA = O
End Function
Sub ZZ_PjDicB()
DicBrw PjDicB(CurPj)
End Sub
Function PjDicB(A As VBProject) As Dictionary
Set PjDicB = DicAy_Mge(CvDicAy(AyMapInto(PjMdAy(A), "MdDicB", EmpDicAy)))
End Function

Function RgOf_FmPj(A As Worksheet) As Range
Set RgOf_FmPj = WsRCRC(A, 1, 1, 3, 1)
End Function

Function RgOf_ToPj(A As Worksheet) As Range
Set RgOf_ToPj = WsRCRC(A, 1, 2, 3, 2)
End Function

Sub SelRg_SetXorEmpty(A As Range)
Dim I
For Each I In A
    
Next
End Sub

Function SrcMthLyItr(A$()) As Collection
Dim O As New Collection

Dim L() As FTIx: L = SrcAllMthFTIxAy(A)
Dim J%
For J = 0 To UB(L)
    O.Add AyWhFTIx(A, L(J))
Next
Set SrcMthLyItr = O
End Function
Sub ZZ_SrcDicA()
DicBrw SrcDicA(CurSrc)
End Sub
Function SrcDicA(A$()) As Dictionary
Dim LyItr As Collection, Ly, MthANm$, O As New Dictionary
Set LyItr = SrcMthLyItr(A)
For Each Ly In LyItr
    MthANm = MthLin_MthANm(Ly(0))
    O.Add MthANm, Ly
Next
Set SrcDicA = O
End Function

Function WsIsPjMgeWs(A As Worksheet) As Boolean
If IsNothing(A) Then Exit Function
If CellOf_ToPjLbl(A).Value <> LblOf_ToPj Then Exit Function
If CellOf_FmPjLbl(A).Value <> LblOf_FmPj Then Exit Function
WsIsPjMgeWs = True
End Function

Function WsOf_PjMge() As Worksheet
Dim O As Worksheet
If WsIsPjMgeWs(CurWs) Then
    Set O = CurWs
Else
    Set O = NewWs
    CellOf_FmPjLbl(O).Value = LblOf_FmPj
    CellOf_ToPjLbl(O).Value = LblOf_ToPj
    CellOf_FmPj(O).Value = "QTool"
    CellOf_ToPj(O).Value = "QVb"
End If
WsVis O
Set WsOf_PjMge = O
End Function

Function ZZDicA() As Dictionary
Set ZZDicA = PjDicA(ZZFmPj)
End Function

Private Function ZZFmPj() As VBProject
Set ZZFmPj = Pj("QTool")
End Function

Function ZZDicB() As Dictionary
Set ZZDicB = PjDicB(ZZToPj)
End Function

Private Function ZZToPj() As VBProject
Set ZZToPj = Pj("QVb")
End Function
Function DicWs(A As Dictionary) As Worksheet
Set DicWs = S1S2Itr_Ws(DicS1S2Ay(A))
End Function
Function S1S2Itr_Ws(A As Collection, Optional Nm1$ = "S1", Optional Nm2$ = "S2") As Worksheet
Set S1S2Itr_Ws = SqWs(S1S2Itr_Sq(A, Nm1, Nm2))
End Function
Function S1S2Itr_Sq(A As Collection, Optional Nm1$ = "S1", Optional Nm2$ = "S2") As Variant()
Dim O(), I, R&
ReDim O(1 To A.Count + 1, 1 To 2)
R = 2
O(1, 1) = Nm1
O(1, 2) = Nm2
For Each I In A
    With CvS1S2(I)
        O(R, 1) = .S1
        O(R, 2) = .S2
        R = R + 1
    End With
Next
S1S2Itr_Sq = O
Stop
End Function
Sub ZZ_DicMaxValSz()
Dim D As Dictionary: Set D = PjDicA(CurPj)
Dim M%: M = DicMaxValSz(D)
Stop
End Sub

Sub ZZ_DoMgePj()
DoMgePj
End Sub


Sub ZZ_DicAB_DryDIF()
DicAB_DryDIF ZZDicA, ZZDicB
End Sub

Sub ZZ_FTPjMSq()
Dim Sq(): Sq = FTPjMSq(ZZFmPj, ZZToPj)
Stop
End Sub

Sub ZZ_MdDicB()
Dim D As Dictionary
Set D = MdDicB(CurMd)
Stop
End Sub

Sub ZZ_PjDicA()
DicBrw PjDicA(CurPj)
End Sub

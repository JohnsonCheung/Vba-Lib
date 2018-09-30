Attribute VB_Name = "IdePjMge"
Option Explicit
Const LblOf_FmPj$ = "Mge from pj"
Const LblOf_ToPj$ = "Mge into pj"

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

Private Sub ZZ_DifOf_BNms()
Dim Act As Collection: Set Act = DifOf_BNms(ZZFmMthDic, ZZToMthDic)
Stop
End Sub
Function DifOf_BNms(FmDicB As Dictionary, ToDicA As Dictionary) As Collection
'See #Dif
'#Dif is for each MthBNm found in FmDicB and also its MthANm is found in ToDicA
'        and the MthB's MthLines is dif and any of the MthA's MthLines
'       (Note.MthANm will have one or more MthLines (due to in differmodule))
Dim BNm, ANm$, MthLinesB$, MthLines, O As New Collection, LinesAy$()
'MthDicB_AssKeysIsBNm FmDicB
For Each BNm In FmDicB.Keys
    ANm = MthBNm_MthANm(BNm)
    If Not ToDicA.Exists(ANm) Then GoTo X
    MthLinesB = FmDicB(BNm)
    LinesAy = ToDicA(ANm)
    If Sz(LinesAy) = 0 Then GoTo X
    For Each MthLines In LinesAy
        If MthLinesB <> MthLines Then
            O.Add BNm
        End If
    Next
X:
Next
Set DifOf_BNms = O
End Function
Private Sub Z_DifOf_Dr()
Dim BNm$: BNm = "G_Tool.AscIsLCase"
Dim Act(): Act = DifOf_Dr(BNm, ZZFmMthDic, ZZToMthDic)
Dim J%, V
'Const C$ = "FmMd ToMd Mth Sel Ty Mdy FmMth "
If Act(0) <> "G_Tool" Then Stop
If Act(1) <> "M_Asc" Then Stop
If Act(2) <> "AscIsLCase" Then Stop
If Not IsMissing(Act(3)) Then Stop
V = Act(4): If V <> "Fun" And V <> "Sub" And V <> "Get" And V <> "Let" And V <> "Set" Then Stop
V = Act(5): If V <> "" And V <> "Prv" And V <> "Frd" Then Stop
For J = 6 To UB(Act)
    If IsEmpty(Act(J)) Then Stop
    If Not IsStr(Act(J)) Then Stop
    If Act(J) = "" Then Stop
Next
End Sub
Private Sub ZZ_MisOf_Dr()
Dim BNm$: BNm = "M_Asc.AscIsLCase"
Dim Act(): Act = MisOf_Dr(BNm, ZZFmMthDic)
Stop
End Sub
Function DifOf_Dr(BNm, FmDicB As Dictionary, ToDicA As Dictionary) As Variant()
Dim Dr(), Ay$(), ToMthLinesAy$(), ANm$
Dr = MisOf_Dr(BNm, FmDicB)
ANm = MthBNm_MthANm(BNm)
'If Not ToDicA.Exists(ANm) Then Stop
ToMthLinesAy = ToDicA(ANm)
DifOf_Dr = AyAddAp(Dr, ToMthLinesAy)
End Function

Function DifOf_Dry(FmDicB As Dictionary, ToDicA As Dictionary) As Variant()
Dim O(), U%, BNm, BNms As Collection, J%
Set BNms = DifOf_BNms(FmDicB, ToDicA)
If BNms.Count = 0 Then Exit Function
ReDim O(BNms.Count - 1)
For Each BNm In BNms
    O(J) = DifOf_Dr(BNm, FmDicB, ToDicA)
    J = J + 1
Next
DifOf_Dry = O
End Function
Private Sub ZZ_PjMgeWs_Bld()
Dim W As Worksheet: Set W = WsOf_PjMge
PjMgeWs_Bld W
WsVis W
End Sub
Sub DoMgePj()
Dim W As Worksheet: Set W = WsOf_PjMge
PjMgeWs_Bld W
'ItrDo PjMgeWs_MthCpyPrms(W), "MthCpyPrm_Cpy"
End Sub

Sub Doc()
'#BNmMIS is Method-B-Nm-of-Missing.
'           Missing means the method is found in FmPj, but not ToPj
'#FmDicB is MthDic-of-MthBNm-zz-MthLines.   It comes from FmPj
'#ToDicA is MthDic-of-MthANm-zz-MthLinesAy. It comes from ToPj
'#ToDicAB is ToDicA and FmDicB
'#ANm is method-a-name, NNN or NNN:YYY
'        If the method is Sub|Fun, just MthNm
'        If the method is Prp    ,      MthNm:MthShtTy
'        It is from ToPj (#ToA)
'        One MthANm will have one or more MthLines
'#BNm is method-b-name, MMM.NNN or MMM.NNN:YYY
'        MdNm.MthNm[:MthShtTy]
'        It is from FmPj (#BFm)
'        One MthBNm will have only one MthLines
'#Missing is for each MthBNm found in FmDicB, but its MthNm is not found in any-method-name-in-ToDicA
'#Dif is for each MthBNm found in FmDicB and also its MthANm is found in ToDicA
'        and the MthB's MthLines is dif and any of the MthA's MthLines
'       (Note.MthANm will have one or more MthLines (due to in differmodule))
End Sub

Private Function FTPjMSq(FmPj As VBProject, ToPj As VBProject) As Variant()
Dim MaxToMthCnt% ' ToPj is using MthNm as key to get all mth (no matter which md and mdy)
                 ' MaxToMthCnt is MaxMthCnt-of-ToPj
Dim ToDicA As Dictionary, FmDicB As Dictionary
Dim Fny$(), A(), B(), Dry()
'Set FmDicB = DicB_RmvMth(PjMthDic(FmPj), "Z__Tst")
'Set ToDicA = DicA_RmvMth(PjDicA(ToPj), "Z__Tst")
MaxToMthCnt = DicMaxValSz(ToDicA)
Fny = MgePjDryFny(MaxToMthCnt)
A = MisOf_Dry(FmDicB, ToDicA)
B = DifOf_Dry(FmDicB, ToDicA)
Dry = AyAddAp(Array(Fny), A, B)
FTPjMSq = DrySq(Dry)
End Function

Function MgePjDryFny(MaxMthCnt%) As String()
Const C$ = "FmMd ToMd Mth Sel Ty Mdy FmMth "
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
MgePjDryFny = SslSy(O)
End Function

Function MisOf_BNms(FmDicB As Dictionary, ToDicA As Dictionary) As Collection
'See #Missing
'#Missing is for each MthBNm found in FmDicB, but its MthNm is not found in any-method-name-in-ToDicA
Dim BNm, O As New Collection
Dim MthNy$(), MthNm$
MthNy = AyMapSy(ToDicA.Keys, "MthANm_MthNm")
For Each BNm In FmDicB.Keys
    MthNm = MthBNm_MthNm(BNm)
    If Not AyHas(MthNy, MthNm) Then
        O.Add BNm
    End If
Next
Set MisOf_BNms = O
End Function

Function MisOf_Dr(BNmMIS, FmDicB As Dictionary) As Variant()
'Const C$ = "FmMd ToMd Mth Sel Ty Mdy FmMth "
Stop
'Dim Ay$(), FmMdNm$, MthNm$
'Dim ToMdNm$, MthShtTy$, ShtMdy$, FmMthLines$
'FmMdNm = MthBNm_MdNm(BNmMIS)
'MthNm = MthBNm_MthNm(BNmMIS)
'ToMdNm = MthProperMdNm(MthNm)
'FmMthLines = FmDicB(BNmMIS): If FmMthLines = "" Then Stop
'ShtMdy = LinShtMdy(FmMthLines)
'MthShtTy = LinMthShtTy(FmMthLines): If MthShtTy = "" Then Stop
'MisOf_Dr = Array(FmMdNm, ToMdNm, MthNm, , MthShtTy, ShtMdy, FmMthLines)
End Function

Function ItrIsEmp(A) As Boolean
ItrIsEmp = A.Count = 0
End Function

Function MisOf_Dry(FmDicB As Dictionary, ToDicA As Dictionary) As Variant()
Dim O(), BNms As Collection, J%, BNm
Set BNms = MisOf_BNms(FmDicB, ToDicA)
If ItrIsEmp(BNms) Then Exit Function
ReDim O(BNms.Count - 1)
J = 0
For Each BNm In BNms
    O(J) = MisOf_Dr(BNm, FmDicB)
    J = J + 1
Next
MisOf_Dry = O
End Function

Function PjMgeSq_MthCpyPrm(Sq, R&, FmPj As VBProject, ToPj As VBProject) As MthCpyPrm
Dim MthNm$, FmMdNm$, ToMdNm$
Dim SrcMth As Mth, ToMd As CodeModule
Dim Mth As Mth
    FmMdNm = Sq(R, 1)
    ToMdNm = Sq(R, 2)
    MthNm = Sq(R, 3)  '<=======
    Set SrcMth = Mth(PjMd(FmPj, FmMdNm), MthNm)
    Set ToMd = PjMd(ToPj, ToMdNm)
Set PjMgeSq_MthCpyPrm = MthCpyPrm(SrcMth, ToMd)
End Function

Function PjMgeSq_MthCpyPrms(Sq, FmPj As VBProject, ToPj As VBProject) As Collection
Dim R&, IsSel, O As Collection
Dim MthNm$, FmMd As CodeModule
For R = 1 To UBound(Sq, 1)
    IsSel = Sq(R, 4)
    If IsSel = "X" Then
        O.Add PjMgeSq_MthCpyPrm(Sq, R, FmPj, ToPj)
    End If
Next
Set PjMgeSq_MthCpyPrms = O
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
SqRg FTPjMSq(FmPj, ToPj), WsRC(A, 3, 1)
End Sub

Private Function PjMgeWs_FmPj(A As Worksheet) As VBProject
Dim P$: P = CellOf_FmPj(A).Value
If Not VbeHasPj(CurVbe, P) Then Exit Function
Set PjMgeWs_FmPj = Pj(P)
End Function

Function PjMgeWs_MthCpyPrms(A As Worksheet) As Collection
Dim Lo As ListObject
If A.ListObjects.Count = 0 Then Exit Function
Set Lo = A.ListObjects(1)
Dim ToPj As VBProject
Dim FmPj As VBProject
    Set ToPj = PjMgeWs_ToPj(A)
    Set FmPj = PjMgeWs_ToPj(A)
SelRg_SetXorEmpty Lo.ListColumns("Sel").DataBodyRange
Set PjMgeWs_MthCpyPrms = PjMgeSq_MthCpyPrms(Lo.DataBodyRange.Value, FmPj, ToPj)
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

Private Property Get ZZToMthDic() As Dictionary
Set ZZToMthDic = PjMthDic(ZZToPj)
End Property

Private Property Get ZZFmMthDic() As Dictionary
Set ZZFmMthDic = PjMthDic(ZZFmPj)
End Property

Private Property Get ZZFmPj() As VBProject
Set ZZFmPj = Pj("QTool")
End Property

Private Property Get ZZToPj() As VBProject
Set ZZToPj = Pj("QVb")
End Property


Private Sub ZZ_DifOf_Dry()
Dim O(): O = DifOf_Dry(ZZFmMthDic, ZZToMthDic)
Stop
End Sub

Private Sub ZZ_DoMgePj()
DoMgePj
End Sub

Private Sub ZZ_FTPjMSq()
Dim Sq(): Sq = FTPjMSq(ZZFmPj, ZZToPj)
Stop
End Sub

Private Sub ZZ_MisOf_Dry()
Dim A(): A = MisOf_Dry(ZZFmMthDic, ZZToMthDic)
Stop
End Sub

Sub Z__Tst()
Stop
End Sub

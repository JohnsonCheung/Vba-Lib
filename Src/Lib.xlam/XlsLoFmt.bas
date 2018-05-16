Attribute VB_Name = "XlsLoFmt"
Option Explicit
Private Type A01 ' Rg Fml
    R As Range
    F As String
End Type
Private Type A02 ' A01()
    A01() As A01
End Type
Private Type NmRslt
    Nm As String
    ErLy() As String
End Type
Private Type BdrRslt
    BdrL As LFCVRslt
    BdrR As LFCVRslt
End Type
Private Type TotRslt
    Avg As LFCVRslt
    Sum As LFCVRslt
    Cnt As LFCVRslt
End Type
Private Type FnyRslt
    Fny() As String
    ErLy() As String
End Type
Private Type LoFmtr
    BdrL As LFCVRslt
    BdrR As LFCVRslt
    Bet  As LFCVRslt
    Cor  As LFCVRslt
    Fml  As LFCVRslt
    Fmt  As LFCVRslt
    Fny() As String
    Hid  As LFCVRslt
    Lbl  As LFCVRslt
    Lvl  As LFCVRslt
    Nm   As String
    Tit  As LFCVRslt
    TAvg As LFCVRslt
    TCnt As LFCVRslt
    TSum As LFCVRslt
    Wdt  As LFCVRslt
End Type
Private Type Brk
    ErLy() As String
    Nm()  As LVF
    Fld() As LVF
    Wdt() As LVF
    Hid() As LVF
    BdrL() As LVF
    BdrR() As LVF
    BdrC() As LVF
    TSum() As LVF
    TAvg() As LVF
    TCnt() As LVF
    Fmt() As LVF
    Lvl() As LVF
    Cor() As LVF
    Lbl() As LFV
    Tit() As LFV
    Fml() As LFV
    Bet() As LFV
End Type
Const C1_Lo$ = "Lo"
Const C1_Fmt$ = "Fmt"
Const C1_Wdt$ = "Wdt"
Const C1_Lvl$ = "Lvl"
Const C1_Cor$ = "Cor"
Const C1_Bdr$ = "Bdr"
Const C1_Tit$ = "Tit"
Const C1_Lbl$ = "Lbl"
Const C1_Fml$ = "Fml"
Const C1_Bet$ = "Bet"
Const C1_Tot$ = "Tot"
Const C2_Tot_Sum$ = "Sum"
Const C2_Tot_Avg$ = "Avg"
Const C2_Tot_Cnt$ = "Cnt"
Const C2_Lo_Nm$ = "Nm"
Const C2_Lo_Fld$ = "Fld"
Const C2_Lo_Hid$ = "Hid"
Const C2_Bdr_L$ = "Left"
Const C2_Bdr_R$ = "Right"
Const C2_Bdr_C$ = "Col"
Sub AAAA()
Tst
End Sub
Sub Tst__Wdt()

End Sub
Private Function ZZWdtRsltLy() As String()
ZZWdtRsltLy = LFCVRslt_OupLy(ZZWdtRslt)
End Function

Private Function ZZWdtRslt() As LFCVRslt
Dim Ly$():       Ly = ZZLoFmtrLy
Dim Brk As Brk: Brk = Ly_Brk(Ly)
Dim W1() As LVF: W1 = Brk.Wdt
Dim Fny() As String: Fny = LVFAy_FnyRslt(Brk.Fld).Fny
Dim W2 As LFCVRslt: W2 = LVFAy_WdtRslt(W1, Fny)
ZZWdtRslt = W2
End Function
Private Function ZZFny() As String()
Dim Ly$():       Ly = ZZLoFmtrLy
Dim Brk As Brk: Brk = Ly_Brk(Ly)
ZZFny = LVFAy_FnyRslt(Brk.Fld).Fny
End Function
Sub Tst()
Dim Ly$(): Ly = ZZLoFmtrLy
Dim A As LoFmtr: A = Ly_Fmtr(Ly)
Stop
DtDmp Fmtr_Dt(A)
End Sub
Private Function LVFAy_NmRslt(Nm() As LVF) As NmRslt
Dim A$: A = LVFAy_Nm(Nm)
Dim ErLy$(): ErLy = AyAddAp( _
    LVFAy_NmErLy_NoNmLin(Nm), _
    LVFAy_NmErLy_ExcessNmLin(Nm), _
    LVFAy_NmErLy_MultiNmInLin(Nm))
With LVFAy_NmRslt
    .ErLy = ErLy
    .Nm = A
End With
End Function
Private Function LVFAy_FnyRslt(Fld() As LVF) As FnyRslt
Dim Fny$(): Fny = LVFAy_Fny(Fld)
With LVFAy_FnyRslt
    .ErLy = LVFAy_LFCVRslt(Fld, Fny, "*Fld").ErLy
    .Fny = Fny
End With
End Function

Private Function LVFAy_Nm$(Nm() As LVF)
Dim Ay$(): Ay = LVFAy_FldLvsAy(Nm)
If AyIsEmp(Ay) Then Exit Function
Dim A$: A = Ay(0)
Dim Ay1$(): Ay1 = LvsSy(A)
If AyIsEmp(Ay1) Then Exit Function
LVFAy_Nm = Ay1(0)
End Function

Private Function LVFAy_NmErLy_NoNmLin(Nm() As LVF) As String()
Dim O$()
If LVF_Sz(Nm) = 0 Then Push O, "There is no line of [Lo Nm *Nm]"
LVFAy_NmErLy_NoNmLin = O
End Function

Private Function LVFAy_NmErLy_ExcessNmLin(Nm() As LVF) As String()
Dim O$(), J%
For J = 0 To LVF_UB(Nm) - 1
    Push O, FmtQQ("Lx(?) Nm-Line, there is also no name line found below.  This line is ignored.", Nm(J).Lx)
Next
LVFAy_NmErLy_ExcessNmLin = O
End Function

Private Function LVFAy_NmErLy_MultiNmInLin(Nm() As LVF) As String()
If LVF_Sz(Nm) = 0 Then Exit Function
Dim O$()
Dim Ay$(): Ay = LVFAy_FldLvsAy(Nm)
Dim U%: U = UB(Ay)
If U = -1 Then Exit Function
Dim A$: A = Ay(U)
If Sz(LvsSy(A)) > 1 Then
    Push O, FmtQQ("Lx(?) has multiple name [?]", Nm(U).Lx, A)
End If
LVFAy_NmErLy_MultiNmInLin = O
End Function

Private Function LFVAy_FmlRslt(A() As LFV, Fny$()) As LFCVRslt
Dim A1 As LFCVRslt: A1 = LFVAy_LFCVRslt(A, Fny, C1_Fml)
Dim J%, ErFldLvs$, OErLy$(), Msg$, Fml$, OAy() As LFCV
For J = 0 To LFCV_UB(A1.Ay)
    Fml = A(J).Val
    ErFldLvs = Fml_ErFldLvs(Fml, Fny)
    If ErFldLvs = "" Then
        LFCV_Push OAy, A1.Ay(J)
    Else
        With A1.Ay(J)
            Msg = FmtQQ("Lx(?) Fld(?) has invalid FldLvs(?)", .Lx, .Fld, ErFldLvs)
        End With
        Push OErLy, Msg
    End If
Next
LFVAy_FmlRslt = LFCVRslt_New(A1, OAy, OErLy)
End Function

Private Function FnyDt(Fny$()) As Dt
FnyDt = AyDt(Fny, "Fld", "InpFny")
End Function

Function LFCVRslt_Dt(A As LFCVRslt) As Dt
With LFCVRslt_Dt
    .Dry = LFCVRslt_Dry(A)
    .Fny = LvsSy("Lx *Itm Fld Cno Val")
    .DtNm = A.Itm
End With
End Function

Sub LFCVRslt_Dmp(A As LFCVRslt)
DsDmp LFCVRslt_Ds(A)
End Sub

Sub LFVAy_FmlRslt__Tst()
Dim Ly$()
Push Ly, "A [B] + [C]"
Push Ly, "B [D] + [E]"
Push Ly, "C [H] * 2"
Push Ly, "G [A]"
Push Ly, "W0 [B]"
Dim Fny$(): Fny = LvsSy("A B C D E F G")
Dim J%
Dim A() As LFV
For J = 0 To UB(Ly)
    With Brk(Ly(J), " ")
        LFV_PushLFV A, (J + 1) * 10, .S1, .S2
    End With
Next
Dim R As LFCVRslt: R = LFVAy_FmlRslt(A, Fny)
LFCVRslt_Dmp R
End Sub


Private Function Fml_ErFldLvs(Fml$, Fny$()) As String
Dim A$(): A = MacroStr_Ny(Fml, ExclBkt:=True, Bkt:="[]")
Dim O$(), J%
For J = 0 To UB(A)
    If Not AyHas(Fny, A(J)) Then
        Push O, A(J)
    End If
Next
Fml_ErFldLvs = JnSpc(O)
End Function

Sub LVFAy_TotRslt__Tst()
Dim Sum() As LVF: LVF_PushLVF Sum, 10, C2_Tot_Sum, "A B E F"
Dim Avg() As LVF: LVF_PushLVF Avg, 20, C2_Tot_Avg, "A B C D"
Dim Cnt() As LVF: LVF_PushLVF Cnt, 30, C2_Tot_Cnt, "A B C D"
Dim Fny$(): Fny = LvsSy("A B C D E")
'AyDmp LVFAy_TotRslt(Sum, Avg, Cnt, Fny)
Stop
End Sub
Function LVFAy_TotRslt(Sum() As LVF, Avg() As LVF, Cnt() As LVF, Fny$()) As TotRslt
Dim A1 As LFCVRslt: A1 = LVFAy_LFCVRslt(Avg, Fny, C2_Tot_Avg)
Dim C1 As LFCVRslt: C1 = LVFAy_LFCVRslt(Cnt, Fny, C2_Tot_Cnt)
Dim S1 As LFCVRslt: S1 = LVFAy_LFCVRslt(Sum, Fny, C2_Tot_Sum)

''-- For B_???C
Dim SumC%():
Dim AvgC%():
Dim CntC%():
    SumC = LFCVAy_CnoAy(S1.Ay)
    AvgC = LFCVAy_CnoAy(A1.Ay)
    CntC = LFCVAy_CnoAy(C1.Ay)
    AvgC = AyMinus(AvgC, SumC)           ' Those avg-column has defined in tot-column, skip
    CntC = AyMinusAp(CntC, SumC, AvgC)
                                         ' Those cnt-column has defined in {tot avg}-column, skip

Dim Ay() As LFCV
Dim ErLy$()

Dim C2 As LFCVRslt
    Ay = LFCVAy_WhByCnoAy(C1.Ay, CntC)
    ErLy = LFCVAy_TotErLy(C1.Ay, AyIntersect(CntC, SumC))
    ErLy = LFCVAy_TotErLy(C1.Ay, AyIntersect(CntC, AvgC))
    C2 = LFCVRslt_New(C1, Ay, ErLy)
    
Dim A2 As LFCVRslt
    Ay = LFCVAy_WhByCnoAy(A1.Ay, AvgC)
    ErLy = LFCVAy_TotErLy(A2.Ay, AyIntersect(AvgC, SumC))
    A2 = LFCVRslt_New(A1, Ay, ErLy)

With LVFAy_TotRslt
    .Avg = A2
    .Cnt = C2
    .Sum = S1
End With
End Function

Private Function LFCVAy_TotErLy(A() As LFCV, CnoAy) As String()
Dim O$(), Msg$, Lx%, F$, Lx2%, J%
For J = 0 To UB(CnoAy)
    Msg = FmtQQ("Lx(?) Fld(?) is also defined in Lx(?)", Lx, F, Lx2)
    Push O, Msg
Next
LFCVAy_TotErLy = O
End Function

Private Function LFCVAy_WhByCnoAy(A() As LFCV, CnoAy%()) As LFCV()
Dim J%, O() As LFCV
For J = 0 To LFCV_UB(A)
    If AyHas(CnoAy, A(J).Cno) Then
        LFCV_Push O, A(J)
    End If
Next
LFCVAy_WhByCnoAy = O
End Function

Sub TotRslt__Tst(Rg As Range)
Static IsInChg As Boolean
If IsInChg Then Exit Sub
IsInChg = True
Dim Ay(), J%, T$
'---------------
Dim Ws As Worksheet:
               Set Ws = RgWs(Rg)
Dim Tit$:         Tit = "Lx *Tot Fld.. Fny Oup"
Dim TitAy$():   TitAy = LvsSy(Tit)
                        AyRgH TitAy, WsRC(Ws, 2, 1) '<== Put Tit
                        WsA1(Ws).Value = "Msg"      '<== Put Msg Tit
Dim MsgRg As Range:
            Set MsgRg = WsRC(Ws, 1, 2)
                        MsgRg.Value = ""           '<== Clear Msg
Dim InpLxRg As Range
Dim InpTotRg As Range
Dim InpFldLvsRg As Range
Dim InpFnyRg As Range
Dim OupRg As Range
          Set InpLxRg = CellVBar(WsRC(Ws, 3, 1), AtLeastOneCell:=True)
         Set InpTotRg = CellVBar(WsRC(Ws, 3, 2), AtLeastOneCell:=True)
      Set InpFldLvsRg = CellVBar(WsRC(Ws, 3, 3), AtLeastOneCell:=True)
         Set InpFnyRg = CellVBar(WsRC(Ws, 3, 4), AtLeastOneCell:=True)
            Set OupRg = CellVBar(WsRC(Ws, 3, 5), AtLeastOneCell:=True)

Dim IsInRg As Boolean:
               IsInRg = CellIsInRgAp(Rg, InpTotRg, InpFldLvsRg)
                        If Not IsInRg Then
                            MsgRg.Value = "Not in range"
                            GoTo X
                        End If
                                                   '<== ShwMsg not in range
Dim InpTot$(): InpTot = VBarSy(InpTotRg)
                        If InpTot(0) = "" Then
                            MsgRg.Value = "1st element of InpLy cannot be empty"
                            GoTo X
                        End If                    '<== ShwMsg if no Input
                        Ay = Array(C2_Tot_Sum, C2_Tot_Avg, C2_Tot_Cnt)
                        For J = 0 To UB(InpTot)
                            T = InpTot(J)
                            If Not AyHas(Ay, T) Then
                                MsgRg.Value = "*Tot column must be one of these [" & JnSpc(Ay) & "]"
                                GoTo X
                            End If
Nxt:
                        Next
Dim InpFldLvs$():
             InpFldLvs = VBarSy(InpFldLvsRg)
                         If InpFldLvs(0) = "" Then
                            MsgRg.Value = "1st element of InpFldLvs cannot be empty"
                            GoTo X
                         End If                    '<== ShwMsg if no Input
Dim InpFny$():
                InpFny = VBarSy(InpFnyRg)
                         If InpFny(0) = "" Then
                            MsgRg.Value = "1st element of InpFld cannot be empty"
                            GoTo X
                         End If                    '<== ShwMsg if no Input

Dim DifSz As Boolean
                 DifSz = Sz(InpTot) <> Sz(InpFldLvs)
                         If DifSz Then
                            MsgRg.Value = "FldLvs & *Tot are dif sz"
                            GoTo X
                         End If

                         AyRgV IntAy_ByU(UB(InpFldLvs)), InpLxRg     '<== Put Lx: 0..

Dim Fny$()
'Run
Dim Sum As LVF
Dim Avg As LVF
Dim Cnt As LVF
                         For J = 0 To UB(InpFldLvs)
                            T = InpTot(J)
                            Select Case T
'                            Case C2_Tot_Sum: Sum  LVF_PushLVF Sum, J, C2_Tot_Sum, InpFldLvs(J))
'                            Case C2_Tot_Avg: Avg  LVF_PushLVF Avg, J, C2_Tot_Avg, InpFldLvs(J))
'                            Case C2_Tot_Cnt: Cnt  LVF_PushLVF Cnt, J, C2_Tot_Cnt, InpFldLvs(J))  '<== Calling
                            Case Else:
                            Stop
                            End Select
                         Next
'Put Rslt
                         OupRg.Clear
                         OupRg.Font.Name = "Courier New"
'Dim RsltLy$():  RsltLy = ZZ51_RsltLy(Sum, Avg, Cnt, Fny)
'                         AyRgH RsltLy, OupRg
X:
    IsInChg = False
End Sub

Function LVFAyOfLeftColRight_BdrRslt(L() As LVF, C() As LVF, R() As LVF, Fny$()) As BdrRslt
Dim L1 As LFCVRslt: L1 = LVFAy_LFCVRslt(L, Fny, "*BdrL")
Dim C1 As LFCVRslt: C1 = LVFAy_LFCVRslt(C, Fny, "*BdrC")
Dim R1 As LFCVRslt: R1 = LVFAy_LFCVRslt(R, Fny, "*BdrR")
Dim C2%(): C2 = LFCVAy_CnoAy(C1.Ay)
Dim L2%(): L2 = LFCVAy_CnoAy(L1.Ay)
Dim R2%(): R2 = LFCVAy_CnoAy(R1.Ay)

Dim O As BdrRslt
'LVFAyOfLeftColRight_BdrRslt = LeftRightLVFAy_ErExistInBothColAndLeft_or_Right(R2, C2, "Right", R1.CnoAy, Fny, C1.CnoAy)
End Function
Function LVFAy_BdrRRslt(R() As LVF, C() As LVF, Fny$()) As LFCVRslt
'Dim L1 As LFCVRslt: L1 = LVFAy_LFCVRslt(L, Fny)
'Dim R1 As LFCVRslt: R1 = LVFAy_LFCVRslt(R, Fny)
'Dim C1 As LFCVRslt: C1 = LVFAy_LFCVRslt(C, Fny)
'Dim C2%(): C2 = C1.CnoAy
'Dim L2%(): L2 = L1.CnoAy
'Dim R2%(): R2 = R1.CnoAy
'Dim L3%(): L3 = AyUniq(AyAddAp(C2, L2))
'Dim R3%(): R3 = AyUniq(AyAddAp(C2, R2))
'Dim E1$(): E1 = LVFAyLeftRighPair_ErLy_OfExistInBothColAndLeft(L2, C2, "Left", L1.CnoAy, Fny, C1.CnoAy)
'Dim E2$(): E2 = LVFAyLeftRighPair_ErLy_OfExistInBothColAndLeft(R2, C2, "Right", R1.CnoAy, Fny, C1.CnoAy)
'Dim ErLy$(): ErLy = AyAddAp(L1.ErLy, R1.ErLy, C1.ErLy, E1, E2)
'With Z115_BdrRslt
'    .BdrL = L3
'    .BdrR = R3
'    .ErLy = ErLy
'End With
End Function
Private Function LVFAyLeftRighPair_ErLy_OfExistInBothColAndLeft(XCnoAy%(), ColCnoAy%(), A$, XLxAy%(), Fny$(), ColLxAy%()) As String()
Dim E%(): E = AyIntersect(XCnoAy, ColCnoAy)
If AyIsEmp(E) Then Exit Function
Dim J%, O$(), Ix%, F$, LxX%, LxC%, ColIx%
For J = 0 To UB(E)
    Ix = E(J)
    F = Fny(Ix)
    LxX = XLxAy(Ix)
    LxC = ColLxAy(ColIx)
    Push O, FmtQQ("Lx(?) is [Bdr ? *Fld..] Fld(?) is also found in Lx(?)-of-[Bdr Col *Fld..]", LxX, A, F, LxC)
Next
LVFAyLeftRighPair_ErLy_OfExistInBothColAndLeft = O
End Function
Private Function Fmtr_Dt(A As LoFmtr) As Dt
Dim O()
With A
    PushAy O, LFCVRslt_Dry(.BdrL)
    PushAy O, LFCVRslt_Dry(.Wdt)
End With
With Fmtr_Dt
    .DtNm = "LoFmtr"
    .Fny = LvsSy("Itm Lx Cno Fld Val")
    .Dry = O
End With
End Function

Private Function LFCVRslt_Dry(A As LFCVRslt) As Variant()
Dim O(), J%
For J = 0 To LFCV_UB(A.Ay)
    With A.Ay(J)
        Push O, Array(A.Itm, .Lx, .Cno, .Fld, .Val)
    End With
Next
LFCVRslt_Dry = O
End Function
Private Function Fmtr_Dry(A As LoFmtr) As Variant()
Dim O()
'With A
'    PushAy O, ZZ41111_Cno(.BdrL, "BdL")
'    PushAy O, ZZ41111_Cno(.BdrR, "BdR")
'    PushAy O, ZZ41111_Bet(.BetA, .BetB, .BetC)
'    PushAy O, LFCVAy_Dry(.Cor, C1_Cor)
'    PushAy O, ZZ41111_Er(.ErLy)
'    PushAy O, ZZ41111_Itm(.Fml, .FmlC, C1_Fml)
'    PushAy O, ZZ41111_Itm(.Fmt, .FmtC, C1_Fmt)
'    PushAy O, ZZ41111_Fny(.Fny)
'    PushAy O, ZZ41111_Cno(.HidC, C2_Lo_Hid)
'    PushAy O, LFCVAy_Dry(.Lbl, C1_Lbl)
'    PushAy O, Array(Array("Nm", , .Nm))
'    PushAy O, ZZ41111_Cno(.TotAvgC, C2_Tot_Avg)
'    PushAy O, ZZ41111_Cno(.TotCntC, C2_Tot_Cnt)
'    PushAy O, ZZ41111_Cno(.TotSumC, C2_Tot_Sum)
'    PushAy O, ZZ41111_Tit(.TitSq)
'    PushAy O, ZZ41111_Itm(.Wdt, .WdtC, C1_Wdt)
'ZZ4111_Dry = O
'Stop
'End With
End Function
Private Sub LFCVAy_FmtBet(A() As LFCV, DtaRg As Range)
Dim J%, F$, F1$, F2$
For J = 0 To LFCV_UB(A)
    With A(J)
        BrkAsg .Val, " ", F1, F2
        F = FmtQQ("=Sum([?]:[?])", F1, F2)
        RgC(DtaRg, .Cno).Formula = F
    End With
Next
End Sub

Sub LoFmtrTp_Brw()
Dim O$()
Push O, "Lo  Nm     *Nm"
Push O, "Lo  Fld    *Fld.."
Push O, "Lo  Hid    *Fld.."
Push O, "Bdr Left   *Fld.."
Push O, "Bdr Right  *Fld.."
Push O, "Bdr Col    *Fld.."
Push O, "Tot Tot    *Fld.."
Push O, "Tot Avg    *Fld.."
Push O, "Tot Cnt    *Fld.."

Push O, "Fmt *Fmt   *Fld.."
Push O, "Wdt *Wdt   *Fld.."
Push O, "Lvl *Lvl   *Fld.."
Push O, "Cor *Cor   *Fld.."

Push O, "Tit *Fld   *Tit"
Push O, "Lbl *Fld   *Lbl"
Push O, "Fml *Fld   *Formula"
Push O, "Bet *Fld   *Fld1 *Fld2"
AyBrw O
End Sub
Private Sub LFCVAy_FmtTit(A() As LFCV, DtaRg As Range)
'Dim R%: R = A.Row - SqNRow(TitSq) - 1
'If R <= 0 Then
'    Debug.Print "LFCVAy_FmtTit: Not enough space to put title at Row=" & R
'    Exit Sub
'End If
'Dim TitRg As Range
'Set TitRg = RgRC(A, R, 1)
'Set TitRg = SqRg(TitSq, TitRg)
'TitRg_Fmt TitRg
End Sub
Private Sub LFCVAy_FmtLbl(A() As LFCV, DtaRg As Range)
Dim J%
'For J = 0 To UB(LblC)
'
'Next
End Sub
Private Sub ZZDmpFny()
AyDmp ZZFny
End Sub
Private Sub ZZResLoFmtrLy()
'Lo Nm ABC
'Lo Fld A B C D E F G
'Lo Hid B C
'Bet A C D
'Bet B D E
'Wdt 10 A B X
'Wdt 20 D C C
'Wdt 30 E F G C
'Fmt #,## A B C
'Fmt #,##.## D E
'Lvl 2 A C
'Bdr Left A
'Bdr Right G
'Bdr Col F
'Tot Sum A B
'Tot Cnt C
'Tot Avg D
'Tit A abc | sdf
'Tit B abc | sdkf | sdfdf
'Cor 12345 A B
'Fml F A + B
'Fml C A * 2
'Lbl A lksd flks dfj
'Lbl B lsdkf lksdf klsdj f
End Sub
Private Sub LoFmt__Tst()
DsWs LoFmtrLy_RsltDs(ZZLoFmtrLy)
End Sub

Sub LoFmt__Tstr(A As Range)
Dim Ws As Worksheet:
Dim Rg As Range
               Set Ws = RgWs(A)
Dim Tit$:         Tit = "Ix InpLoFmtrLy"
Dim TitAy$():   TitAy = LvsSy(Tit)
                        AyRgH TitAy, WsRC(Ws, 2, 1) '<== Put Tit
                        WsA1(Ws).Value = "Msg"      '<== Put Msg Tit
Dim MsgRg As Range:
            Set MsgRg = WsRC(Ws, 1, 2)
                        MsgRg.Value = ""           '<== Clear Msg
Dim InpLyRg As Range:
               Set Rg = WsRC(Ws, 3, 4)
          Set InpLyRg = CellVBar(Rg)
Dim IsInRg As Boolean:
               IsInRg = CellIsInRgAp(A, InpLyRg)
                        If Not IsInRg Then
                            MsgRg.Value = "Not in range"
                            Exit Sub
                        End If
                                                          '<== ShwMsg not in range
Dim InpLy$():    InpLy = VBarSy(InpLyRg)
                         If InpLy(0) = "" Then
                            MsgRg.Value = "1st element of InpLy cannot be empty"
                            Exit Sub
                         End If                           '<== ShwMsg if no Input

Dim LnoRg As Range:
             Set LnoRg = RgRC(InpLyRg, 1, 0)
                         CellClrDown LnoRg
                         CellFillSeqDown LnoRg, Sz(InpLy)    '<== Put Lno: 0..
Dim RsltDs As Ds
                RsltDs = LoFmtrLy_RsltDs(InpLy)
Dim OupRg As Range
                Set Rg = WsRC(Ws, 3, 4)
             Set OupRg = CellVBar(Rg)
                         OupRg.Clear
                         'AyRgV(RsltLy, OupRg).Font = "Courier New"             '<== Put Rslt
End Sub

Private Function LoFmtrLy_RsltDs(LoFmtrLy$()) As Ds
Dim F As LoFmtr: F = Ly_Fmtr(LoFmtrLy)
Dim D1 As Dt: D1 = AyDt(Fmtr_Ly(F), "Formatted Ly", "Oup")
Dim D2 As Dt: D2 = Fmtr_Dt(F)
Dim D3 As Dt: 'D3 = ErLy_Dt(F.ErLy)
Dim O As Ds
O.DsNm = "LoFmtr Rslt"
DsAddDt O, D1
DsAddDt O, D2
DsAddDt O, D3
LoFmtrLy_RsltDs = O
End Function
Private Function Ly_Brk(Ly$()) As Brk
Dim O As Brk, A$, B$, C$, J%
With O
    For J = 0 To UB(Ly)
        LinAsgTTRst Ly(J), A, B, C
        Select Case A
        Case C1_Lo
            Select Case B
            Case "Hid": LVF_PushLVF .Hid, J, C2_Lo_Hid, C
            Case "Fld": LVF_PushLVF .Fld, J, C2_Lo_Fld, C
            Case "Nm":  LVF_PushLVF .Nm, J, C2_Lo_Nm, C
            End Select
        Case C1_Bdr
            Select Case B
            Case C2_Bdr_L: LVF_PushLVF .BdrL, J, C2_Bdr_L, C
            Case C2_Bdr_R: LVF_PushLVF .BdrR, J, C2_Bdr_R, C
            Case C2_Bdr_C: LVF_PushLVF .BdrC, J, C2_Bdr_C, C
            Case Else: Push .ErLy, FmtQQ("Lx(?) T2(?) should be [Left Right Col]", J, B)
            End Select
        Case C1_Tot
            Select Case B
            Case C2_Tot_Sum:  LVF_PushLVF .TSum, J, C2_Tot_Sum, C
            Case C2_Tot_Avg:  LVF_PushLVF .TAvg, J, C2_Tot_Avg, C
            Case C2_Tot_Cnt:  LVF_PushLVF .TCnt, J, C2_Tot_Cnt, C
            Case Else:  Push .ErLy, FmtQQ("Lx(?) T2(?) should be [Tot Avg Cnt]", J, B)
            End Select
        Case C1_Fmt: LVF_PushLVF .Fmt, J, B, C
        Case C1_Wdt: LVF_PushLVF .Wdt, J, B, C
        Case C1_Lvl: LVF_PushLVF .Lvl, J, B, C
        Case C1_Cor: LVF_PushLVF .Cor, J, B, C
        Case C1_Tit: LFV_PushLFV .Tit, J, B, C
        Case C1_Lbl: LFV_PushLFV .Lbl, J, B, C
        Case C1_Fml: LFV_PushLFV .Fml, J, B, C
        Case C1_Bet: LFV_PushLFV .Bet, J, B, C
        Case Else
            Push .ErLy, FmtQQ("Lx(?) T1(?) should be [Lo Wdt Lbl ...]", J, A)
        End Select
    Next
End With
Ly_Brk = O
End Function
Private Function Ly_Fmtr(Ly$()) As LoFmtr
Dim Brk As Brk: Brk = Ly_Brk(Ly)
With Brk
    Dim RNm As NmRslt:    RNm = LVFAy_NmRslt(.Nm)
    Dim RFld As FnyRslt: RFld = LVFAy_FnyRslt(.Fld)
    Dim Fny$(): Fny = RFld.Fny
    Dim RBdr As BdrRslt: RBdr = LVFAyOfLeftColRight_BdrRslt(.BdrL, .BdrC, .BdrR, Fny)
    Dim RWdt As LFCVRslt: RWdt = LVFAy_WdtRslt(.Wdt, Fny)
    Dim RHid As LFCVRslt: RHid = LVFAy_LFCVRslt(.Hid, Fny, C1_Lo)
    Dim RTot As TotRslt:  RTot = LVFAy_TotRslt(.TSum, .TAvg, .TCnt, Fny)
    Dim RFmt As LFCVRslt: RFmt = LVFAy_LFCVRslt(.Fmt, Fny, C1_Fmt)
    Dim RLvl As LFCVRslt: RLvl = LVFAy_LvlRslt(.Lvl, Fny)
    Dim RCor As LFCVRslt: RCor = LVFAy_CorRslt(.Cor, Fny)
    Dim RLbl As LFCVRslt: RLbl = LFVAy_LFCVRslt(.Lbl, Fny, C1_Lbl)
    Dim RTit As LFCVRslt: RTit = LFVAy_LFCVRslt(.Tit, Fny, C1_Tit)
    Dim RFml As LFCVRslt: RFml = LFVAy_FmlRslt(.Fml, Fny)
    Dim RBet As LFCVRslt: RBet = LFVAy_BetRslt(.Bet, Fny)
End With

Dim O As LoFmtr
With O
    .Nm = RNm.Nm
    .Fny = RFld.Fny
    .BdrL = RBdr.BdrL
    .BdrR = RBdr.BdrR
    .Bet = RBet
    .Cor = RCor
    .Fml = RFml
    .Fmt = RFmt
    .Hid = RHid
    .Lbl = RLbl
    .Lvl = RLvl
    .Tit = RTit
    .TSum = RTot.Sum
    .TAvg = RTot.Avg
    .TCnt = RTot.Cnt
    .Wdt = RWdt
End With
Ly_Fmtr = O
End Function
Sub LoFmt(Lo As ListObject, LoFmtrLy$())
Dim R As Range
Dim J%, A() As LFCV
Dim F As LoFmtr: F = Ly_Fmtr(LoFmtrLy)
A = F.BdrL.Ay: For J = 0 To LFCV_UB(A): RgBdrLeft RgC(R, A(J).Cno):               Next
A = F.BdrR.Ay: For J = 0 To LFCV_UB(A): RgBdrRight RgC(R, A(J).Cno):              Next
A = F.Fml.Ay:  For J = 0 To LFCV_UB(A): RgC(R, A(J).Cno).Formula = A(J).Val:      Next
A = F.Fmt.Ay:  For J = 0 To LFCV_UB(A): RgC(R, A(J).Cno).NumberFormat = A(J).Val: Next
A = F.Wdt.Ay:  For J = 0 To LFCV_UB(A): RgC(R, A(J).Cno).ColumnWidth = A(J).Val:  Next
A = F.Hid.Ay:  For J = 0 To LFCV_UB(A): RgC(R, A(J).Cno).Hidden = True:           Next
A = F.Lvl.Ay:  For J = 0 To LFCV_UB(A): RgC(R, A(J).Cno).OutlineLevel = A(J).Val: Next
A = F.TSum.Ay: For J = 0 To LFCV_UB(A): LoColNm_SetTot Lo, A(J).Fld:              Next
A = F.TAvg.Ay: For J = 0 To LFCV_UB(A): LoColNm_SetAvg Lo, A(J).Fld:              Next
A = F.TCnt.Ay: For J = 0 To LFCV_UB(A): LoColNm_SetCnt Lo, A(J).Fld:              Next
A = F.Cor.Ay:  For J = 0 To LFCV_UB(A): RgC(R, A(J).Cno).Interior.Color = A(J).Val: Next
LFCVAy_FmtTit F.Tit.Ay, R
LFCVAy_FmtLbl F.Lbl.Ay, R
End Sub

Private Function ZZLoFmtrLy() As String()
ZZLoFmtrLy = MdResLy(Md("XlsLoFmt"), "LoFmtrLy")
End Function

Function Fmtr_Ly(A As LoFmtr, Optional InclErLy As Boolean) As String()
Dim O$()
With A
    Push O, FmtQQ("? ? ?", C1_Lo, C2_Lo_Nm, .Nm)
    Push O, FmtQQ("? ? ?", C1_Lo, C2_Lo_Fld, JnSpc(.Fny))
    PushAy O, LFCVAy_ValFldLy(.Hid.Ay, C1_Lo)
    
    PushAy O, LFCVAy_FldValLy(.Fml.Ay, C1_Fml)

    PushAy O, LFCVAyOfBdrPair_ValFldLy(.BdrL.Ay, .BdrR.Ay)
    PushAy O, LFCVAy_ValFldLy(.TSum.Ay, C1_Tot)
    PushAy O, LFCVAy_ValFldLy(.TAvg.Ay, C1_Tot)
    PushAy O, LFCVAy_ValFldLy(.TCnt.Ay, C1_Tot)
    PushAy O, LFCVAy_ValFldLy(.Fmt.Ay, C1_Fmt)
    PushAy O, LFCVAy_ValFldLy(.Wdt.Ay, C1_Wdt)
    PushAy O, LFCVAy_ValFldLy(.Lvl.Ay, C1_Lvl)
    PushAy O, LFCVAy_ValFldLy(.Cor.Ay, C1_Cor)
    PushAy O, LFCVAy_FldValLy(.Tit.Ay, C1_Tit)
    PushAy O, LFCVAy_FldValLy(.Lbl.Ay, C1_Lbl)
    PushAy O, LFCVAy_FldValLy(.Bet.Ay, C1_Bet)
'    If InclErLy Then PushAy O, AyAddPfx(.ErLy, "-- ")
End With
Fmtr_Ly = O
End Function
Private Function LFCVAy_ValFldLy(A() As LFCV, Itm$) As String()

End Function

Private Function LFCVAy_FldValLy(A() As LFCV, Itm$) As String()

End Function

Private Function LFCVAyOfBdrPair_ValFldLy(L() As LFCV, R() As LFCV) As String()
'Dim J%, ColFny$(), RightFny$(), LeftFny$()
'Dim C%(): C = AyIntersect(L, R)
'Dim L1%(): L1 = AyMinus(L, C)
'Dim R1%(): R1 = AyMinus(R, C)
'For J = 0 To UB(C)
'    Push O, T1 & " Col " & Fny(C(J))
'Next
'For J = 0 To UB(L1)
'    Push O, T1 & " Left " & Fny(L1(J))
'Next
'For J = 0 To UB(R1)
'    Push O, T1 & " Right " & Fny(R1(J))
'Next
End Function

Function LVFAy_WdtRslt(A() As LVF, Fny$()) As LFCVRslt
Dim A1 As LFCVRslt:   A1 = LVFAy_LFCVRslt(A, Fny, C1_Wdt)
LVFAy_WdtRslt = LFCVRslt_VdtValBet(A1, 2, 100)
End Function

Private Sub LFCVRslt_WdtRslt__Tst()
Dim Ly$()
Push Ly, "20 A B C D"
Push Ly, "30 D E F"
Push Ly, "330 X Y Z"
Push Ly, "W0 I H"
Dim J%, Inp() As LVF
For J = 0 To UB(Ly)
    With Brk(Ly(J), " ")
        LVF_PushLVF Inp, J * 10, .S1, .S2
    End With
Next
Dim Fny$(): Fny = LvsSy("A B C D E X")
Dim R As LFCVRslt: R = LVFAy_WdtRslt(Inp, Fny)
LFCVRslt_Dmp R
Stop
End Sub

Private Function LFCV_Sz%(A() As LFCV)
On Error Resume Next
LFCV_Sz = UBound(A) + 1
End Function

Private Function LFCV_UB%(A() As LFCV)
LFCV_UB = LFCV_Sz(A) - 1
End Function

Private Sub LFCV_Push(O() As LFCV, A As LFCV)
Dim N%: N = LFCV_Sz(O)
ReDim Preserve O(N)
O(N) = A
End Sub

Private Function LVFAy_LvlRslt(A() As LVF, Fny$()) As LFCVRslt
Dim A1 As LFCVRslt: A1 = LVFAy_LFCVRslt(A, Fny, C1_Lvl)
LVFAy_LvlRslt = LFCVRslt_VdtValBet(A1, 2, 6)
End Function

Private Function LVFAy_CorRslt(A() As LVF, Fny$()) As LFCVRslt
Dim A1 As LFCVRslt: A1 = LVFAy_LFCVRslt(A, Fny, C1_Cor)
LVFAy_CorRslt = LFCVRslt_VdtValIsLng(A1)
End Function
Private Function BetVal_ErMsg$(A$, Fny$())
Dim Ay$(): Ay = LvsSy(A)
If Sz(Ay) <> 2 Then BetVal_ErMsg = "it has " & Sz(Ay) & " terms, not 2": Exit Function
Dim Ay1$()
If Not AyHas(Fny, Ay(0)) Then Push Ay1, Ay(0)
If Not AyHas(Fny, Ay(1)) Then Push Ay1, Ay(1)
If Sz(Ay1) = 0 Then Exit Function
BetVal_ErMsg = FmtQQ("Term ? are invalid", JnQSqBktSpc(Ay1))
End Function
Private Function LFVAy_BetRslt(A() As LFV, Fny$()) As LFCVRslt
Dim A1 As LFCVRslt: A1 = LFVAy_LFCVRslt(A, Fny, C1_Bet)
Dim OErLy$(), X$, Msg$
Dim J%, OAy() As LFCV
For J = 0 To LFCV_UB(A1.Ay)
    With A1.Ay(J)
        X = BetVal_ErMsg(.Val, Fny)
        If X = "" Then
            Msg = FmtQQ("Lx(?) Fld(?) should have 2 valid fields, but now [?]", .Lx, .Fld, X)
            Push OErLy, Msg
        Else
            LFCV_Push OAy, A1.Ay(J)
        End If
    End With
Next
LFVAy_BetRslt = LFCVRslt_New(A1, OAy, OErLy)
End Function
Private Function ZZRsltDmp(A As LFCVRslt)
LFCVRslt_Dmp A
End Function
Private Sub ZZ_Wdt()
AyDmp ZZWdtIOLy
End Sub
Private Function ZZWdtIOLy() As String()
ZZWdtIOLy = AyAddAp(ZZWdtInpLy, ZZFny, ZZWdtRsltLy)
End Function
Private Function ZZWdtInpLy() As String()
ZZWdtInpLy = DrsLy(LVFAy_Drs(ZZWdtInp))
End Function

Private Function ZZWdtInp() As LVF()
Dim Brk As Brk: Brk = Ly_Brk(ZZLoFmtrLy)
ZZWdtInp = Brk.Wdt
End Function
Private Sub ZZDmpWdtInp()
LVFAy_Dmp ZZWdtInp
End Sub

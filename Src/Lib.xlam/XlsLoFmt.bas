Attribute VB_Name = "XlsLoFmt"
Option Explicit
Public Const C1_Lo$ = "Lo"
Public Const C1_Fmt$ = "Fmt"
Public Const C1_Wdt$ = "Wdt"
Public Const C1_Lvl$ = "Lvl"
Public Const C1_Cor$ = "Cor"
Public Const C1_Bdr$ = "Bdr"
Public Const C1_Tit$ = "Tit"
Public Const C1_Lbl$ = "Lbl"
Public Const C1_Fml$ = "Fml"
Public Const C1_Bet$ = "Bet"
Public Const C1_Tot$ = "Tot"
Public Const C2_Tot_Sum$ = "Sum"
Public Const C2_Tot_Avg$ = "Avg"
Public Const C2_Tot_Cnt$ = "Cnt"
Public Const C2_Lo_Nam$ = "Nam"
Public Const C2_Lo_Fld$ = "Fld"
Public Const C2_Lo_Hid$ = "Hid"
Public Const C2_Bdr_L$ = "Left"
Public Const C2_Bdr_R$ = "Right"
Public Const C2_Bdr_C$ = "Col"
'Private A_Ly$() ' LoFmtrLy
'Private A_Lo As ListObject
'Private Type BdrB: Left() As LVF: Right() As LVF: Column() As LVF: End Type
'Private Type TotB: Sum() As LVF: Avg() As LVF: Cnt() As LVF: End Type
'
'Private Type Vdted: O() As LCFV: ErLy() As String: Ly() As String: End Type
'Private Type TotO: Sum As Vdted: Avg As Vdted: Cnt As Vdted: End Type
'Private Type BdrO: LCno() As Integer: RCno() As Integer: End Type
'Private Type BdrV: O As BdrO: ErLy() As String: Ly() As String: End Type
'Private Type TotV: O As TotO: ErLy() As String: Ly() As String: End Type
'Private Type FnyV: O() As String: ErLy() As String: Ly() As String: End Type
'Private Type NmV:  O As String: ErLy() As String: Ly() As String: End Type
'
'Private Type Brk
'    FnyB() As LVF
'    BdrB   As BdrB
'    BetB() As LFV
'    CorB() As LVF
'    FmlB() As LFV
'    FmtB() As LVF
'    HidB() As LVF
'    LblB() As LFV
'    LvlB() As LVF
'    NmB()  As LVF
'    TitB() As LFV
'    TotB   As TotB
'    WdtB() As LVF
'    ErLy() As String
'End Type
'Private Type Vdt
'    FnyV As FnyV
'    BdrV As BdrV
'    BetV As Vdted
'    CorV As Vdted
'    FmlV As Vdted
'    FmtV As Vdted
'    HidV As Vdted
'    LblV As Vdted
'    LvlV As Vdted
'    NmV  As NmV
'    TitV As Vdted
'    TotV As TotV
'    WdtV As Vdted
'    ErLy() As String
'End Type
'Private Type Oup
'    BdrO   As BdrO
'    BetO() As LCFV
'    CorO() As LCFV
'    FmlO() As LCFV
'    FmtO() As LCFV
'    HidO() As LCFV
'    LblO() As LCFV
'    TitO() As LCFV
'    TotO   As TotO
'    WdtO() As LCFV
'End Type
'
'Private Function ZBdrV(A As BdrB, Fny$()) As BdrV
'With ZBdrV
'
'End With
'End Function
'
'Private Function Ly(A As Vdt) As String()
'Dim O$()
'With A
'    O = AyAddAp(.NmV.Ly, .FnyV.Ly, .HidV.Ly, .FmtV.Ly, .FmlV.Ly, .CorV.Ly, _
'        .TotV.O.Avg.Ly, _
'        .TotV.O.Sum.Ly, _
'        .TotV.O.Cnt.Ly, _
'        .TitV.Ly, .LblV.Ly, .LvlV.Ly, .ErLy)
'End With
'End Function
'Private Function ZVdt(A As Brk) As Vdt
'Dim O As Vdt
'With O
'    Dim F$()
'    .NmV = ZNmV(A.NmB)
'    .BetV = ZBetV(A.BetB, F)
'    .BdrV = ZBdrV(A.BdrB, F)
'    .FnyV = ZFnyV(A.FnyB)
'    .HidV = ZHidV(A.HidB, F)
'    .FmtV = ZFmtV(A.FmtB, F)
'    .FmlV = ZFmlV(A.FmlB, F)
'    .LblV = ZLblV(A.LblB, F)
'    .LvlV = ZLvlV(A.LvlB, F)
'    .TitV = ZTitV(A.TitB, F)
'    .TotV = ZTotV(A.TotB, F)
'    .CorV = ZCorV(A.CorB, F)
'    .ErLy = A.ErLy
'End With
'ZVdt = O
'End Function
'Private Function ZHidV(HidB() As LVF, Fny$()) As Vdted
'Dim A As LCFVRslt: A = LVFAy_LCFVRslt(HidB, Fny)
''ZHidV.O = A
'End Function
'Private Function ZFmtV(FmtB() As LVF, Fny$()) As Vdted
'
'End Function
'Private Function ZLvlV(LvlB() As LVF, Fny$()) As Vdted
'
'End Function
'Private Function ZLblV(LblB() As LFV, Fny$()) As Vdted
'
'End Function
'Private Function ZTitV(TitB() As LFV, Fny$()) As Vdted
'
'End Function
'Private Function ZCorV(CorB() As LVF, Fny$()) As Vdted
'
'End Function
'
'Private Function ZBrk(Ly$()) As Brk
'Dim O As Brk
'ZBrk = O
'End Function
'Function ErLy() As String()
''With ZVdt(Brk)
''    ErLy = A.AddAp(.Er, .Nm.Er, .Fny.Er, .Hid.Er, .Fmt.Er, .Fml.Er, .Cor.Er, .Tot.Er, .Tit.Er, .Lbl.Er, .Lvl.Er).ErLy
''End With
'End Function
'
'Function FmtrLy() As String()
''With ZVdt
''    FmtrLy = AyAddAp(.Nm, .Fny, .Hid.Ly, .Fmt.Ly, .Fml.Ly, .Cor.Ly, .Tot.Ly, .Tit.Ly, .Lbl.Ly, .Lvl.Ly)
''End With
'End Function
'Sub ZDoFmtBdr(A As BdrO, DtaRg As Range)
'
'End Sub
'Sub ZDoFmtBet(BetO() As LCFV, DtaRg As Range)
'
'End Sub
'Sub Fmt(A As Oup, DtaRg As Range)
'With A
'    ZDoFmtBdr .BdrO, DtaRg
'    ZDoFmtBet .BetO, DtaRg
'    'ZodFmtCor .CorO, DtaRg
'End With
'End Sub
'
'Private Function ZOup(A As Vdt) As Oup
'Dim O As Oup
'With O
'    .BdrO = A.BdrV.O
'    .BetO = A.BetV.O
'    .CorO = A.CorV.O
'    .FmlO = A.CorV.O
'    .FmtO = A.FmtV.O
'    .HidO = A.HidV.O
'    .LblO = A.LblV.O
'    .TitO = A.TitV.O
'    .TotO = A.TotV.O
'    .WdtO = A.WdtV.O
'End With
'ZOup = O
'End Function
'Private Function ZBdrO()
'
'End Function
'
'Private Function ZTotV(A As TotB, Fny$()) As TotV
''Dim A1 As LCFVRslt: A1 = LVFAy_LCFVRslt(Avg, Fny, C2_Tot_Avg)
''Dim C1 As LCFVRslt: C1 = LVFAy_LCFVRslt(Cnt, Fny, C2_Tot_Cnt)
''Dim S1 As LCFVRslt: S1 = LVFAy_LCFVRslt(Sum, Fny, C2_Tot_Sum)
''
''''-- For B_???C
''Dim SumC%():
''Dim AvgC%():
''Dim CntC%():
''    SumC = LCFVAy_CnoAy(S1.Ay)
''    AvgC = LCFVAy_CnoAy(A1.Ay)
''    CntC = LCFVAy_CnoAy(C1.Ay)
''    AvgC = AyMinus(AvgC, SumC)           ' Those avg-column has defined in tot-column, skip
''    CntC = AyMinusAp(CntC, SumC, AvgC)
''                                         ' Those cnt-column has defined in {tot avg}-column, skip
''
''Dim Ay() As LCFV
''Dim ErLy$()
''
''Dim C2 As LCFVRslt
''    Ay = LCFVAy_WhByCnoAy(C1.Ay, CntC)
''    ErLy = ZErLy_Tot_DupDef(C1.Ay, AyIntersect(CntC, SumC))
''    ErLy = ZErLy_Tot_DupDef(C1.Ay, AyIntersect(CntC, AvgC))
''    C2 = LCFVRslt_Add_Ay_Er(C1, Ay, ErLy)
''
''Dim A2 As LCFVRslt
''    Ay = LCFVAy_WhByCnoAy(A1.Ay, AvgC)
''    ErLy = ZErLy_Tot_DupDef(A2.Ay, AyIntersect(AvgC, SumC))
''    A2 = LCFVRslt_Add_Ay_Er(A1, Ay, ErLy)
''
''With ZTotV
''    .Avg = A2
''    .Cnt = C2
''    .Sum = S1
''End With
'End Function
'
'Private Function ZWdtV() '(A As WdtB, Fny$()) As WdtV
'End Function
'Private Function LCFVAy_S1S2Ay(A() As LCFV) As S1S2()
'Dim O() As S1S2, J%
'For J = 0 To LCFV_UB(A)
'    With A(J)
'        S1S2_Push O, NewS1S2(.F, .V)
'    End With
'Next
'LCFVAy_S1S2Ay = O
'End Function
'
'
'Private Sub ZZZ_Fml()
'Dim A() As LFV
'    Dim Ly$()
'    Push Ly, "A [B] + [C]"
'    Push Ly, "B [D] + [E]"
'    Push Ly, "C [H] * 2"
'    Push Ly, "G [A]"
'    Push Ly, "W0 [B]"
'    Dim Fny$(): Fny = LvsSy("A B C D E F G")
'    Dim J%
'    For J = 0 To UB(Ly)
'        With Brk(Ly(J), " ")
'            LFV_Push3 A, (J + 1) * 10, .S1, .S2
'        End With
'    Next
''Dim R As LCFVRslt: R = ZFmlV(A, Fny)
''LCFVRslt_Dmp R
'End Sub
'Private Function ZZLoFmtIOLy() As String()
'
'End Function
'Private Sub ZZZ_LoFmt()
'AyDmp ZZLoFmtIOLy
'End Sub
'
'Private Sub ZZ_TotV()
''Dim B As TotB
''Dim Sum() As LVF: LVF_Push3 Sum, 10, C2_Tot_Sum, "A B E F"
''Dim Avg() As LVF: LVF_Push3 Avg, 20, C2_Tot_Avg, "A B C D"
''Dim Cnt() As LVF: LVF_Push3 Cnt, 30, C2_Tot_Cnt, "A B C D"
''With B
''    .Sum = Sum
''    .Avg = Avg
''    .Cnt = Cnt
''End With
''Dim Fny$(): Fny = LvsSy("A B C D E")
''AyDmp ZTotVLy(ZTotV(B, Fny))
''Stop
'End Sub
'
'Private Function ZErLy_Bet_TermEr$(A$, Fny$())
'Dim Ay$(): Ay = LvsSy(A)
'If Sz(Ay) <> 2 Then ZErLy_Bet_TermEr = "it has " & Sz(Ay) & " terms, not 2": Exit Function
'Dim Ay1$()
'If Not AyHas(Fny, Ay(0)) Then Push Ay1, Ay(0)
'If Not AyHas(Fny, Ay(1)) Then Push Ay1, Ay(1)
'If Sz(Ay1) = 0 Then Exit Function
'ZErLy_Bet_TermEr = FmtQQ("Term ? are invalid", JnQSqBktSpc(Ay1))
'End Function
'
'Private Function ZFEFldLvs$(Fml$, Fny$()) ' 'Fml Er Fld Lvs :
'Dim A$(): A = MacroStr_Ny(Fml, ExclBkt:=True, Bkt:="[]")
'Dim O$(), J%
'For J = 0 To UB(A)
'    If Not AyHas(Fny, A(J)) Then
'        Push O, A(J)
'    End If
'Next
'ZFEFldLvs = JnSpc(O)
'End Function
'Private Function ZVBdrX() '(L() As LCFV, R() As LCFV) As BdrLCR
''Dim C1() As LCFV: C1 = LCFV_Intersect(L, R)
''Dim L1() As LCFV: L1 = LCFV_Minus(L, C1)
''Dim R1() As LCFV: R1 = LCFV_Minus(R, C1)
''With ZVBdrX
''    .C = C1
''    .L = L1
''    .R = R1
''End With
'End Function
'
'Private Function ZVdtDry() '(A As Vdt) As Variant()
'Dim O()
''With A
''    With ZVBdrX(.BdrL.Ay, .BdrR.Ay)
''        PushAy O, LCFVRslt_Dry(.L, C2_Bdr_L)
''        PushAy O, LCFVRslt_Dry(.C, C2_Bdr_C)
''        PushAy O, LCFVRslt_Dry(.R, C2_Bdr_R)
''    End With
''    PushAy O, LCFVRslt_Dry(.Bet)
''    PushAy O, LCFVRslt_Dry(.Cor)
''    PushAy O, LCFVRslt_Dry(.ErLy)
''    PushAy O, ZZ41111_Itm(.Fml, .FmlC, C1_Fml)
''    PushAy O, ZZ41111_Itm(.Fmt, .FmtC, C1_Fmt)
''    PushAy O, ZZ41111_Fny(.Fny)
''    PushAy O, ZZ41111_Cno(.HidC, C2_Lo_Hid)
''    PushAy O, LCFVAy_Dry(.Lbl, C1_Lbl)
''    PushAy O, Array(Array("Nm", , .Nm))
''    PushAy O, ZZ41111_Cno(.TotAvgC, C2_Tot_Avg)
''    PushAy O, ZZ41111_Cno(.TotCntC, C2_Tot_Cnt)
''    PushAy O, ZZ41111_Cno(.TotSumC, C2_Tot_Sum)
''    PushAy O, ZZ41111_Tit(.TitSq)
''    PushAy O, ZZ41111_Itm(.Wdt, .WdtC, C1_Wdt)
''ZZ4111_Dry = O
''Stop
''End With
'End Function
'
'Private Function LCFVAy2_ValFldLy(L() As LCFV, R() As LCFV) As String()
''Dim J%, ColFny$(), RightFny$(), LeftFny$()
''Dim C%(): C = AyIntersect(L, R)
''Dim L1%(): L1 = AyMinus(L, C)
''Dim R1%(): R1 = AyMinus(R, C)
''For J = 0 To UB(C)
''    Push O, T1 & " Col " & Fny(C(J))
''Next
''For J = 0 To UB(L1)
''    Push O, T1 & " Left " & Fny(L1(J))
''Next
''For J = 0 To UB(R1)
''    Push O, T1 & " Right " & Fny(R1(J))
''Next
'End Function
'
'Private Function LCFVAy_FldValLy(A() As LCFV, Itm$) As String()
'End Function
'
'
'Private Function ZErLy_Tot_DupDef(A() As LCFV, CnoAy) As String()
'Dim O$(), Msg$, Lx%, F$, Lx2%, J%
'For J = 0 To UB(CnoAy)
'    Msg = FmtQQ("Lx(?) Fld(?) is also defined in Lx(?)", Lx, F, Lx2)
'    Push O, Msg
'Next
'ZErLy_Tot_DupDef = O
'End Function
'
'
'Private Function ZBetV(BetB() As LFV, Fny$()) As Vdted
''Dim A1 As LCFVRslt: ' A1 = LFVAy_LCFVRslt(A, Fny, C1_Bet)
''Dim OErLy$(), X$, Msg$
''Dim J%, OAy() As LCFV
''For J = 0 To LCFV_UB(A1.Ay)
''    With A1.Ay(J)
''        X = ZErLy_Bet_TermEr(.V, Fny)
''        If X = "" Then
''            Msg = FmtQQ("Lx(?) Fld(?) should have 2 valid fields, but now [?]", .Lx, .F, X)
''            Push OErLy, Msg
''        Else
''            LCFV_Push OAy, A1.Ay(J)
''        End If
''    End With
''Next
''ZBetV = LCFVRslt_Add_Ay_Er(A1, OAy, OErLy)
'End Function
'
'Private Function ZFmlV(Fml() As LFV, Fny$()) As Vdted
''Dim A1 As LCFVRslt: ' A1 = LFVAy_LCFVRslt(A, Fny, C1_Fml)
''Dim J%, ErFldLvs$, OErLy$(), Msg$, Fml$, OAy() As LCFV
''For J = 0 To LCFV_UB(A1.Ay)
''    'Fml = A(J).Val
''    ErFldLvs = ZFEFldLvs(Fml, Fny)
''    If ErFldLvs = "" Then
''        LCFV_Push OAy, A1.Ay(J)
''    Else
''        With A1.Ay(J)
''            Msg = FmtQQ("Lx(?) Fld(?) has invalid FldLvs(?)", .Lx, .F, ErFldLvs)
''        End With
''        Push OErLy, Msg
''    End If
''Next
''ZFmlV = LCFVRslt_Add_Ay_Er(A1, OAy, OErLy)
'End Function
'
'Private Function ZErLy_Bdr(XCnoAy%(), ColCnoAy%(), A$, XLxAy%(), Fny$(), ColLxAy%()) As String()
'Dim E%(): E = AyIntersect(XCnoAy, ColCnoAy)
'If AyIsEmp(E) Then Exit Function
'Dim J%, O$(), Ix%, F$, LxX%, LxC%, ColIx%
'For J = 0 To UB(E)
'    Ix = E(J)
'    F = Fny(Ix)
'    LxX = XLxAy(Ix)
'    LxC = ColLxAy(ColIx)
'    Push O, FmtQQ("Lx(?) is [Bdr ? *Fld..] Fld(?) is also found in Lx(?)-of-[Bdr Col *Fld..]", LxX, A, F, LxC)
'Next
'ZErLy_Bdr = O
'End Function
'
'Private Function ZFnyV(FnyB() As LVF) As FnyV
''Dim Fny$(): Fny = LVFAy_Fny(Fld)
''With ZVFny
'    '.ErLy = LVFAy_LCFVRslt(Fld, Fny, "*Fld").ErLy
'    '.Fny = Fny
''End With
'End Function
'
'
'Private Function ZNm$(Nm() As LVF)
'Dim Ay$(): Ay = LVFAy_FldLvsAy(Nm)
'If AyIsEmp(Ay) Then Exit Function
'Dim A$: A = Ay(0)
'Dim Ay1$(): Ay1 = LvsSy(A)
'If AyIsEmp(Ay1) Then Exit Function
'ZNm = Ay1(0)
'End Function
'
'Private Function ZErLy_Nm_ExcessLin(Nm() As LVF) As String()
'Dim O$(), J%
'For J = 0 To LVF_UB(Nm) - 1
'    Push O, FmtQQ("Lx(?) Nm-Line, there is also no name line found below.  This line is ignored.", Nm(J).Lx)
'Next
'ZErLy_Nm_ExcessLin = O
'End Function
'
'Private Function ZErLy_Nm_MultiLIn(Nm() As LVF) As String()
'If LVF_Sz(Nm) = 0 Then Exit Function
'Dim O$()
'Dim Ay$(): Ay = LVFAy_FldLvsAy(Nm)
'Dim U%: U = UB(Ay)
'If U = -1 Then Exit Function
'Dim A$: A = Ay(U)
'If Sz(LvsSy(A)) > 1 Then
'    Push O, FmtQQ("Lx(?) has multiple name [?]", Nm(U).Lx, A)
'End If
'ZErLy_Nm_MultiLIn = O
'End Function
'
'Private Function ZErLy_Nm_NoLin(Nm() As LVF) As String()
'Dim O$()
'If LVF_Sz(Nm) = 0 Then Push O, "There is no line of [Lo Nm *Nm]"
'ZErLy_Nm_NoLin = O
'End Function
'
'Private Function ZNmV(NmB() As LVF) As NmV
''Dim A$: A = ZNm(Nm)
''Dim ErLy$(): ErLy = AyAddAp( _
''    ZErLy_Nm_NoLin(Nm), _
''    ZErLy_Nm_ExcessLin(Nm), _
''    ZErLy_Nm_MultiLIn(Nm))
''With ZVNm
''    .ErLy = ErLy
''    .Nm = A
''End With
'End Function
'
'
'Private Sub ZZDmpFny()
'AyDmp ZZIFny
'End Sub
'
'Private Function ZZVdt() ' As Vdt
''ZZVdt = ZVdt(ZZBrk)
'End Function
'
'Private Function ZZIFny() As String()
'Dim Fny$(): 'Fny = ZFnyV(ZZBrk.Fny).O.Fny
'ZZIFny = ApSy("Inp-Fny", Sy(Fny).IxLy)
'End Function
'
'Private Function ZZLoFmtrLy() As String()
''ZZLoFmtrLy = MdResLy(Md("XlsLoFmt"), "LoFmtrLy")
'End Function
'
'Private Sub ZZResLoFmtrLy()
''Lo Nm ABC
''Lo Fld A B C D E F G
''Lo Hid B C
''Bet A C D
''Bet B D E
''Wdt 10 A B X
''Wdt 20 D C C
''Wdt 30 E F G C
''Fmt #,## A B C
''Fmt #,##.## D E
''Lvl 2 A C
''Bdr Left A
''Bdr Right G
''Bdr Col F
''Tot Sum A B
''Tot Cnt C
''Tot Avg D
''Tit A abc | sdf
''Tit B abc | sdkf | sdfdf
''Cor 12345 A B
''Fml F A + B
''Fml C A * 2
''Lbl A lksd flks dfj
''Lbl B lsdkf lksdf klsdj f
'End Sub
'
'Private Sub ZZZ_LCFVRslt_WdtRslt()
''Dim Ly$()
''Push Ly, "20 A B C D"
''Push Ly, "30 D E F"
''Push Ly, "330 X Y Z"
''Push Ly, "W0 I H"
''Dim J%, Inp() As LVF
''For J = 0 To UB(Ly)
''    With Brk(Ly(J), " ")
''        LVF_Push3 Inp, J * 10, .S1, .S2
''    End With
''Next
''Dim Fny$(): Fny = LvsSy("A B C D E X")
''Dim R As LCFVRslt: R = ZVWdt(Inp, Fny)
''LCFVRslt_Dmp R
'XlsLoFmt.ZZZ_Fml
'End Sub
'
'Sub XX()
'Dim A As New Dictionary
'A.Add 1, 1
'Dim X
'Dim AA As Worksheet
'
'For Each X In A
'    Stop
'Next
'End Sub
Sub AAA()
Dim A As New VBA.Collection
A.Add "skdl", "A"
A.Add "lskdfj, 1"
Debug.Print A.Count
Debug.Print A("A")
Debug.Print A("1")
Stop
End Sub

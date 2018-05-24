VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoFmtr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type Pair: ValVy As Variant: CnoAy() As Integer: End Type
Private Type BdrV: Ok() As String: Er As Er: BdrL() As Integer: BdrR() As Integer: End Type
Private Type BetV: Ok() As String: Er() As String: BetC() As Integer: BetA() As String: BetB() As String: End Type
Private Type CorV: Ok() As String: Er() As String: CorC() As Integer: Cor()  As Long: End Type
Private Type FmlV: Ok() As String: Er() As String: FmlC() As Integer: Fml() As String: End Type
Private Type FmtV: Ok() As String: Er() As String: FmtC() As Integer: Fmt() As String: End Type
Private Type FnyV: Ok() As String: Er() As String: Fny() As String: End Type
Private Type HidV: Ok() As String: Er() As String: HidC() As Integer: End Type
Private Type LblV: Ok() As String: Er() As String: LblC() As Integer: Lbl() As String: End Type
Private Type LvlV: Ok() As String: Er() As String: LvlC() As Integer: Lvl() As Byte: End Type
Private Type TitV: Ok() As String: Er() As String: TitC() As Integer: Tit() As String: End Type
Private Type TotV: Ok() As String: Er() As String: Sum() As String: Avg() As String: Cnt() As String: End Type
Private Type WdtV: Ok() As String: Er() As String: WdtC() As Integer: Wdt() As Integer: End Type
Private Type BdrI: Left As New P3Itm: Right As New P3Itm: Col As New P3Itm: End Type
Private Type TotI: Sum  As New P3Itm: Avg   As New P3Itm: Cnt As New P3Itm: End Type


Property Get ToStr$()
End Property

Property Get Validate() As LoFmtrRslt
Dim Ok As New LoFmtr
Dim Er As New Er
Dim F As New FmtWs
'Dim BdrI As BdrI
'Dim BetI As P3Itm: BetI = P3Itm(C1_Bet)
'Dim CorI As New P3Itm
'Dim FmlI As New P3Itm
'Dim FmtI As New P3Itm
'Dim FnyI As New P3Itm
'Dim HidI As New P3Itm
'Dim LblI As New P3Itm
'Dim LvlI As New P3Itm
'Dim NamI As New P3Itm
'Dim TitI As New P3Itm
'Dim TotI As TotI
'Dim WdtI As New P3Itm
'Dim Itm As P3Itm
'Dim J%, A$, B$, C$
'For J = 0 To UB(Ly)
'    LinAsgTTRst Ly(J), A, B, C
'    Select Case A
'    Case C1_Lo
'        Select Case B
'            Case C2_Lo_Hid: Set Itm = HidI
'            Case C2_Lo_Fld: Set Itm = FnyI
'            Case C2_Lo_Nam: Set Itm = NamI
'        End Select
'    Case C1_Bdr
'        Select Case B
'            Case C2_Bdr_L: Set Itm = BdrI.Left
'            Case C2_Bdr_R: Set Itm = BdrI.Right
'            Case C2_Bdr_C: Set Itm = BdrI.Col
'            Case Else: Push Er, FmtQQ("Lx(?) T2(?) should be [Left Right Col]", J, B)
'        End Select
'    Case C1_Tot
'        Select Case B
'            Case C2_Tot_Sum: Set Itm = TotI.Sum
'            Case C2_Tot_Avg: Set Itm = TotI.Avg
'            Case C2_Tot_Cnt: Set Itm = TotI.Cnt
'            Case Else:  Push Er, FmtQQ("Lx(?) T2(?) should be [Tot Avg Cnt]", J, B)
'        End Select
'    Case C1_Fmt: Set Itm = FmtI
'    Case C1_Fml: Set Itm = FmlI
'    Case C1_Bet: Set Itm = BetI
'    Case C1_Tit: Set Itm = TitI
'    Case C1_Lbl: Set Itm = LblI
'    Case C1_Wdt: Set Itm = WdtI
'    Case C1_Lvl: Set Itm = LvlI
'    Case C1_Cor: Set Itm = CorI
'    Case Else
'        Push Er, FmtQQ("Lx(?) T1(?) should be [Lo Wdt Lbl ...]", J, A)
'    End Select
'    Itm.PushLin J, B, C
'Next
'Dim FnyV As P3FnyRslt: Set FnyV = FnyI.VdtAsFny
'Dim Fny$(): Fny = FnyV.Fny
'Dim BdrV As BdrV: BdrV = ZBdrV(BdrI, Fny)
'Dim BetV As BetV: BetV = ZBetV(BetI, Fny)
'Dim CorV As CorV: CorV = ZCorV(CorI, Fny)
'Dim FmlV As FmlV: FmlV = ZFmlV(FmlI, Fny)
'Dim FmtV As FmtV: FmtV = ZFmtV(FmtI, Fny)
'Dim HidV As HidV: HidV = ZHidV(HidI, Fny)
'Dim LblV As LblV: LblV = ZLblV(LblI, Fny)
'Dim LvlV As LvlV: LvlV = ZLvlV(LvlI, Fny)
'Dim NamV As P3NmRslt: Set NamV = ZNamV(NamI)
'Dim TitV As TitV: TitV = ZTitV(TitI, Fny)
'Dim TotV As TotV: TotV = ZTotV(TotI, Fny)
'Dim WdtV As WdtV: WdtV = ZWdtV(WdtI, Fny)
'B_Er = AyAddAp(Er, BdrV.Er, BetV.Er, CorV.Er, FmlV.Er, HidV.Er, LblV.Er, LvlV.Er, TitV.Er, TotV.Er, WdtV.Er)
'B_Ok = AyAddAp(BdrV.OK, BetV.OK, CorV.OK, FmlV.OK, HidV.OK, LblV.OK, LvlV.OK, TitV.OK, TotV.OK, WdtV.OK)
'With BdrV: F.SetBdr .BdrL, .BdrR:        End With
'With BetV: F.SetBet .BetC, .BetA, .BetB: End With
'With CorV: F.SetCor .CorC, .Cor:         End With
'With FmlV: F.SetFml .FmlC, .Fml:         End With
'With FmtV: F.SetFmt .FmtC, .Fmt:         End With
'With HidV: F.SetHid .HidC:               End With
'With LblV: F.SetLbl .LblC, .Lbl:         End With
'With LvlV: F.SetLvl .LvlC, .Lvl:         End With
'With TitV: F.SetTit .TitC, .Tit:         End With
'With TotV: F.SetTot .Sum, .Avg, .Cnt:    End With
'With WdtV: F.SetWdt .WdtC, .Wdt:         End With
Dim O As New LoFmtrRslt
Set O.Er = Er
Set O.LoFmtr = Ok
Set O.FmtWs = F
Set Validate = O
End Property
'Private Function ZBdrV(BdrI As BdrI, Fny$()) As BdrV
'
'End Function
'Private Function ZBetV(BetI As P3Itm, Fny$()) As BetV
'
'End Function
'Private Function ZCorV(CorI As P3Itm, Fny$()) As CorV
'
'End Function
'Private Function ZFmlEr(FmlI As P3Itm) As P3ItmRslt
'
'End Function
'Private Function ZFmlV(FmlI As P3Itm, Fny$()) As FmlV
'Dim V As P3ItmRslt: Set V = FmlI.Vdt(Fny)
'With ZFmlEr(V.Lins)
'    Dim Ay() As LVF: ' Ay =
'    Dim E1$(): E1 = .Er
'End With
'Dim Fml$()
'Dim FmlC%()
'With ZFmlV
'    .Er = AyAddAp(V.Er, E1)
'    .OK = OK
'    .Fml = Fml
'    .FmlC = FmlC
'End With
'End Function
'
'Private Function ZFmtV(FmtI As P3Itm, Fny$()) As FmtV
'End Function
'
'Private Function ZHidV(HidI As P3Itm, Fny$()) As HidV
'End Function
'
'Private Function ZLblV(LblI As P3Itm, Fny$()) As LblV
'End Function
'
'Private Function ZLvlV(LvlI As P3Itm, Fny$()) As LvlV
'End Function
'
'Private Function ZNamV(NamI As P3Itm) As P3NmRslt
'Set ZNamV = NamI.VdtAsNm(C1_Lo, C2_Lo_Nam)
'End Function
'Private Function ZTitV(TitI As P3Itm, Fny$()) As TitV
'
'End Function
'Private Function ZTotV(TotI As TotI, Fny$()) As TotV
'
'End Function
'Private Function ZWdtV(WdtI As P3Itm, Fny$()) As WdtV
'Dim R As P3ItmRslt: Set R = WdtI.VdtAsNumBet(Fny, 2, 100)
'Dim A As P3Itm: Set A = R.Lins
'Dim P As Pair: P = ZPair(A)
'With ZWdtV
'    .Wdt = P.ValVy
'    .WdtC = P.CnoAy
'    .Er = R.Er
'    .OK = A.Ly
'End With
'End Function
'Private Function ZPair(A As P3Itm) As Pair
'
'End Function
'Private Function ZWdtOk(Ay() As P3Lin) As String()
'End Function
'Friend Sub Tst()
'AyDmp ZZIOLy
'End Sub
'Private Function ZIOLy() As String()
'Dim O$()
'PushItmAy O, "Inp::LoFmtrLy--------------------------", A_Ly
'PushItmAy O, "Oup1::Ok::LoFmtrLy---------------------", B_Ok
'PushItmAy O, "Oup2::Er::Ly---------------------------", B_Er
'PushItmAy O, "Oup3::FmtWs::Show----------------------", B_FmtWs.Show
'Push O, "*End"
'ZIOLy = O
'End Function
'Private Function ZZIOLy() As String()
'Init ZZLoFmtrLy
'ZZIOLy = ZIOLy
'End Function
'Private Function ZZLoFmtrLy() As String()
'ZZLoFmtrLy = MdResLy(IdeMd.Md("LoFmtr"), "LoFmtrLy")
'End Function
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

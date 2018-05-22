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
Private A_Ly$()
Private B_Ok$()
Private B_Er$()
Private B_FmtWs As FmtWs
Private Type LVFRslt: Ay() As LVF: Er() As String: End Type
Private Type BdrV: Ok() As String: Er() As String: BdrL() As Integer: BdrR() As Integer: End Type
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
Private Type BdrB: Left As New P3: Right As New P3: Col As New P3: End Type
Private Type TotB: Sum  As New P3: Avg   As New P3: Cnt As New P3: End Type
Sub BrwTp()
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

Property Get Ok() As String()
Ok = B_Ok
End Property
Property Get Er() As String()
Er = B_Er
End Property
Friend Property Get FmtWs() As FmtWs
Set FmtWs = B_FmtWs
End Property
Friend Function Init(Ly$()) As LoFmtr
A_Ly = Ly
Dim Er$(), Ok$(), F As New FmtWs
Dim BdrB As BdrB
Dim BetB As New P3
Dim CorB As New P3
Dim FmlB As New P3
Dim FmtB As New P3
Dim FnyB As New P3
Dim HidB As New P3
Dim LblB As New P3
Dim LvlB As New P3
Dim NamB As New P3
Dim TitB As New P3
Dim TotB As TotB
Dim WdtB As New P3

Dim J%, A$, B$, C$
For J = 0 To UB(Ly)
    LinAsgTTRst Ly(J), A, B, C
    Select Case A
    Case C1_Lo
        Select Case B
            Case "Hid": HidB.PushLVF J, C2_Lo_Hid, C
            Case "Fld": FnyB.PushLVF J, C2_Lo_Fld, C
            Case "Nm":  NamB.PushLVF J, C2_Lo_Nm, C
        End Select
    Case C1_Bdr
        Select Case B
            Case C2_Bdr_L: BdrB.Left.PushLVF J, C2_Bdr_L, C
            Case C2_Bdr_R: BdrB.Right.PushLVF J, C2_Bdr_R, C
            Case C2_Bdr_C: BdrB.Col.PushLVF J, C2_Bdr_C, C
        Case Else: Push Er, FmtQQ("Lx(?) T2(?) should be [Left Right Col]", J, B)
        End Select
    Case C1_Tot
        Select Case B
            Case C2_Tot_Sum: TotB.Sum.PushLVF J, C2_Tot_Sum, C
            Case C2_Tot_Avg: TotB.Avg.PushLVF J, C2_Tot_Avg, C
            Case C2_Tot_Cnt: TotB.Cnt.PushLVF J, C2_Tot_Cnt, C
        Case Else:  Push Er, FmtQQ("Lx(?) T2(?) should be [Tot Avg Cnt]", J, B)
        End Select
        Case C1_Fmt: FmtB.PushLVF J, B, C
        Case C1_Wdt: WdtB.PushLVF J, B, C
        Case C1_Lvl: LvlB.PushLVF J, B, C
        Case C1_Cor: CorB.PushLVF J, B, C
        Case C1_Tit: TitB.PushLFV J, B, C
        Case C1_Lbl: LblB.PushLFV J, B, C
        Case C1_Fml: FmlB.PushLVF J, B, C
        Case C1_Bet: BetB.PushLFV J, B, C
    Case Else
        Push Er, FmtQQ("Lx(?) T1(?) should be [Lo Wdt Lbl ...]", J, A)
    End Select
Next
Dim FnyV As P3FnyRslt: Set FnyV = FnyB.VdtAsFny
Dim Fny$(): Fny = FnyV.Fny
Dim BdrV As BdrV: BdrV = ZBdrV(BdrB, Fny)
Dim BetV As BetV: BetV = ZBetV(BetB, Fny)
Dim CorV As CorV: CorV = ZCorV(CorB, Fny)
Dim FmlV As FmlV: FmlV = ZFmlV(FmlB, Fny)
Dim FmtV As FmtV: FmtV = ZFmtV(FmtB, Fny)
Dim HidV As HidV: HidV = ZHidV(HidB, Fny)
Dim LblV As LblV: LblV = ZLblV(LblB, Fny)
Dim LvlV As LvlV: LvlV = ZLvlV(LvlB, Fny)
Dim NamV As P3NmRslt: Set NamV = ZNamV(NamB)
Dim TitV As TitV: TitV = ZTitV(TitB, Fny)
Dim TotV As TotV: TotV = ZTotV(TotB, Fny)
Dim WdtV As WdtV: WdtV = ZWdtV(WdtB, Fny)
B_Er = AyAddAp(Er, BdrV.Er, BetV.Er, CorV.Er, FmlV.Er, HidV.Er, LblV.Er, LvlV.Er, TitV.Er, TotV.Er, WdtV.Er)
B_Ok = AyAddAp(BdrV.Ok, BetV.Ok, CorV.Ok, FmlV.Ok, HidV.Ok, LblV.Ok, LvlV.Ok, TitV.Ok, TotV.Ok, WdtV.Ok)
With BdrV: F.SetBdr .BdrL, .BdrR:        End With
With BetV: F.SetBet .BetC, .BetA, .BetB: End With
With CorV: F.SetCor .CorC, .Cor:         End With
With FmlV: F.SetFml .FmlC, .Fml:         End With
With FmtV: F.SetFmt .FmtC, .Fmt:         End With
With HidV: F.SetHid .HidC:               End With
With LblV: F.SetLbl .LblC, .Lbl:         End With
With LvlV: F.SetLvl .LvlC, .Lvl:         End With
With TitV: F.SetTit .TitC, .Tit:         End With
With TotV: F.SetTot .Sum, .Avg, .Cnt:    End With
With WdtV: F.SetWdt .WdtC, .Wdt:         End With
Set B_FmtWs = F
End Function
Private Function ZBdrV(BdrB As BdrB, Fny$()) As BdrV

End Function
Private Function ZBetV(BetB As P3, Fny$()) As BetV

End Function
Private Function ZCorV(CorB As P3, Fny$()) As CorV

End Function
Private Function ZFmlEr(A() As P3LCFV) As LVFRslt

End Function
Private Function ZFmlV(FmlB As P3, Fny$()) As FmlV
Dim V As P3LCFVRslt: Set V = FmlB.Vdt(Fny)
With ZFmlEr(V.LCFVAy)
    Dim Ay() As LVF: Ay = .Ay
    Dim E1$(): E1 = .Er
End With
Dim Fml$()
Dim FmlC%()
With ZFmlV
    .Er = AyAddAp(V.Er, E1)
    .Ok = Ok
    .Fml = Fml
    .FmlC = FmlC
End With
End Function

Private Function ZFmtV(FmtB As P3, Fny$()) As FmtV
End Function

Private Function ZHidV(HidB As P3, Fny$()) As HidV
End Function

Private Function ZLblV(LblB As P3, Fny$()) As LblV
End Function

Private Function ZLvlV(LvlB As P3, Fny$()) As LvlV
End Function

Private Function ZNamV(NamB As P3) As P3NmRslt
Set ZNamV = NamB.VdtAsNm(C1_Lo, C2_Lo_Nm)
End Function
Private Function ZTitV(TitB As P3, Fny$()) As TitV

End Function
Private Function ZTotV(TotB As TotB, Fny$()) As TotV

End Function
Private Function ZWdtV(WdtB As P3, Fny$()) As WdtV
Dim V As P3LCFVRslt: Set V = WdtB.VdtAsNumBet(Fny, 2, 100)
Dim Ay() As P3LCFV: Ay = V.LCFVAy
Dim Wdt%(), WdtC%()
With ZWdtV
    .Wdt = Wdt
    .WdtC = WdtC
    .Er = V.Er
    .Ok = ZWdtOk(Ay)
End With
End Function
Private Function ZWdtOk(Ay() As P3LCFV) As String()
End Function
Friend Sub Tst()
AyDmp ZZIOLy
End Sub
Private Function ZIOLy() As String()
Dim O$()
PushItmAy O, "Inp::LoFmtrLy--------------------------", A_Ly
PushItmAy O, "Oup1::Ok::LoFmtrLy---------------------", B_Ok
PushItmAy O, "Oup2::Er::Ly---------------------------", B_Er
PushItmAy O, "Oup3::FmtWs::Show----------------------", B_FmtWs.Show
Push O, "*End"
ZIOLy = O
End Function
Private Function ZZIOLy() As String()
Init ZZLoFmtrLy
ZZIOLy = ZIOLy
End Function
Private Function ZZLoFmtrLy() As String()
ZZLoFmtrLy = MdResLy(IdeMd.Md("LoFmtr"), "LoFmtrLy")
End Function
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


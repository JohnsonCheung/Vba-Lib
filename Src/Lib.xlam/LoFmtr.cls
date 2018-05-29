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
Private B_Ly$()
Private Type CnoStr
    Cno() As Integer
    Str() As String
End Type
Private Type CnoLng
    Cno() As Integer
    Lng() As Long
End Type
Private Type Bet
    BetC() As Integer
    BetA() As String
    BetB() As String
End Type
Function Init(Ly$()) As LoFmtr
B_Ly = Ly
Set Init = Me
End Function
Private Function VdtFml(A As LABCsRslt, Fny$()) As LABCsRslt

End Function
Private Function VdtTot(A As P3Tot, Fny$()) As P3TotRslt

End Function
Private Function VdtBdr(A As P3Bdr, Fny$()) As LABCsRslt

End Function
Property Get ToStr$()
ToStr = "LoFmtr(" & vbCrLf & JnCrLf(B_Ly) & vbCrLf & "LoFmtr)"
End Property

Property Get Validate() As LoFmtrRslt
Dim O As New LoFmtrRslt
Dim I As P3Inp
Dim J%, A$, B$, C$
Dim Er$()
For J = 0 To UB(B_Ly)
    Lin(B_Ly(J)).AsgTTRst A, B, C
    Select Case A
    Case C1_Lo
        Select Case B
            Case C2_Lo_Hid: I.VF.Hid.AddLBC J, B, C
            Case C2_Lo_Fld: I.VF.Fny.AddLBC J, B, C
            Case C2_Lo_Nam: I.VF.Nam.AddLBC J, B, C
        End Select
    Case C1_Bdr
        Select Case B
            Case C2_Bdr_L: I.VF.Bdr.BdrL.AddLBC J, B, C
            Case C2_Bdr_R: I.VF.Bdr.BdrR.AddLBC J, B, C
            Case C2_Bdr_C: I.VF.Bdr.BdrC.AddLBC J, B, C
            Case Else: Push Er, FmtQQ("Lx(?) T2(?) should be [Left Right Col]", J, B)
        End Select
    Case C1_Tot
        Select Case B
            Case C2_Tot_Sum: I.VF.Tot.Sum.AddLBC J, B, C
            Case C2_Tot_Avg: I.VF.Tot.Avg.AddLBC J, B, C
            Case C2_Tot_Cnt: I.VF.Tot.Cnt.AddLBC J, B, C
            Case Else:  Push Er, FmtQQ("Lx(?) T2(?) should be [Tot Avg Cnt]", J, B)
        End Select
    Case C1_Fmt: I.VF.Fmt.AddLBC J, B, C
    Case C1_Fml: I.FV.Fml.AddLBC J, B, C
    Case C1_Bet: I.FV.Bet.AddLBC J, B, C
    Case C1_Tit: I.FV.Tit.AddLBC J, B, C
    Case C1_Lbl: I.FV.Tit.AddLBC J, B, C
    Case C1_Wdt: I.VF.Wdt.AddLBC J, B, C
    Case C1_Lvl: I.VF.Lvl.AddLBC J, B, C
    Case C1_Cor: I.VF.Cor.AddLBC J, B, C
    Case Else
        Push Er, FmtQQ("Lx(?) T1(?) should be [Lo Wdt Lbl ...]", J, A)
    End Select
Next
Dim NamR As NmRslt: Set NamR = I.VF.Nam.ValidateAsNm
Dim FnyR As FnyRslt: Set FnyR = I.VF.Fny.ValidateAsFny
Dim Fny$(): Fny = FnyR.Fny
Dim BdrR  As P3BdrRslt:  Set BdrR = VdtBdr(I.VF.Bdr, Fny)
Dim BetR  As LABCsRslt:  Set BetR = I.FV.Bet.ValidateAsFldVal(Fny)
Dim CorR  As LABCsRslt:  Set CorR = I.VF.Cor.ValidateAsFldLngVal(Fny)
Dim FmlR  As LABCsRslt:  Set FmlR = VdtFml(I.FV.Fml.ValidateAsFldVal(Fny), Fny)
Dim FmtR  As LABCsRslt:  Set FmtR = I.VF.Fmt.ValidateAsFldVal(Fny)
Dim HidR  As LABCsRslt:  Set HidR = I.VF.Hid.ValidateAsFldVal(Fny)
Dim LblR  As LABCsRslt:  Set LblR = I.FV.Lbl.ValidateAsFldVal(Fny)
Dim LvlR  As LABCsRslt:  Set LvlR = I.VF.Lvl.ValidateAsBetNum(Fny, 2, 6)
Dim TitR  As LABCsRslt:  Set TitR = I.FV.Tit.ValidateAsFldVal(Fny)
Dim TotR  As P3TotRslt:  Set TotR = VdtTot(I.VF.Tot, Fny)
Dim WdtR  As LABCsRslt:  Set WdtR = I.VF.Wdt.ValidateAsFldVal(Fny)
O.Er.PushAp BdrR.Er, BetR.Er, CorR.Er, FmlR.Er, HidR.Er, LblR.Er, LvlR.Er, TitR.Er, TotR.Er, WdtR.Er
Dim Ok$
Ok = StrAp_Lines( _
    NamR.Lin, _
    FnyR.Lin, _
    LvlR.LABCs.Ly, _
    BdrR.LABCs.Ly, _
    BetR.LABCs.Ly, _
    CorR.LABCs.Ly, _
    FmlR.LABCs.Ly, _
    FmtR.LABCs.Ly, _
    HidR.LABCs.Ly, _
    LblR.LABCs.Ly, _
    LblR.LABCs.Ly, _
    TitR.LABCs.Ly, _
    TotR.LABCs.Ly, _
    WdtR.LABCs.Ly)

Dim XBdrL%(), XBdrR%()
Dim XBetC%(), XBetA$(), XBetB$()
Dim XCorC%(), XCor&()
Dim XFmlC%(), XFml$()
Dim XFmtC%(), XFmt$()
Dim XLvlC%(), XLvl() As Byte
Dim XTitC%(), XTit$()
Dim XHidC%()
Dim XLblC%(), XLbl$()
Dim XTSum$(), XTAvg$(), XTCnt$()
Dim XWdtC%(), XWdt%()

With O.FmtWs
    .SetWdt1 WdtR.LABCs.CnoAyWdtAy(Fny)

    .SetBdr XBdrL, XBdrR
    .SetBet XBetC, XBetA, XBetB
    .SetCor XCorC, XCor
    .SetFml XFmlC, XFml
    .SetFmt XFmtC, XFmt
    .SetHid XHidC
    .SetLbl XLblC, XLbl
    .SetLvl XLvlC, XLvl
    .SetTit XTitC, XTit
    .SetTot XTSum, XTAvg, XTCnt
    .SetWdt XWdtC, XWdt
End With
Set Validate = O
End Property

'Private Function ZFmlV(FmlI As LABCs, Fny$()) As FmlV
'Dim V As LABCsRslt: Set V = FmlI.Vdt(Fny)
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
Private Function ZBet(A As LABCs) As Bet
Stop
End Function
Private Function ZCnoLng(A As LABCs) As CnoLng
Stop
End Function
Private Function ZCnoStr1(A As LABCs) As CnoStr
Stop
End Function


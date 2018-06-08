VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FmtWs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit
'Private B_BdrCnoAy%()
'Private B_HidCnoAy%()
'Private B_SumCny$(), B_AvgCny$(), B_CntCny$() 'Cny - Col Name Array
'Private B_Bet As New CnoVals
'Private B_Cor As New CnoVals
'Private B_Fml As New CnoVals
'Private B_Fmt As New CnoVals
'Private B_Lbl As New CnoVals
'Private B_Tit As New CnoVals
'Private B_Lvl As New CnoVals
'Private B_Wdt As New CnoVals
'
'Friend Sub SetBdr(Cno%())
'B_BdrCnoAy = Cno
'End Sub
'
'Property Get Bdr() As IntAyObj: Set Bdr = IntAyObj(B_BdrCnoAy): End Property
'Property Get Hid() As IntAyObj: Set Hid = IntAyObj(B_HidCnoAy): End Property
'Property Get Sum() As SyObj:    Set Sum = SyObj(B_SumCny): End Property
'Property Get Avg() As SyObj:    Set Avg = SyObj(B_AvgCny): End Property
'Property Get Cnt() As SyObj:    Set Cnt = SyObj(B_CntCny): End Property
'Property Get Bet() As CnoVals:   Set Bet = B_Bet: End Property
'Property Get Cor() As CnoVals:   Set Cor = B_Cor: End Property
'Property Get Fml() As CnoVals:   Set Fml = B_Fml: End Property
'Property Get Fmt() As CnoVals:   Set Fmt = B_Fmt: End Property
'Property Get Lbl() As CnoVals:   Set Lbl = B_Lbl: End Property
'Property Get Tit() As CnoVals:   Set Tit = B_Tit: End Property
'Property Get Lvl() As CnoVals:   Set Lvl = B_Lvl: End Property
'Property Get Wdt() As CnoVals:   Set Wdt = B_Wdt: End Property
'
'Friend Sub SetCor(A As CnoVals)
'Set B_Cor = A
'End Sub
'
'Friend Sub SetWdt(A As CnoVals)
'Set B_Wdt = A
'End Sub
'
'Property Get ToStr$()
'Dim O$()
'ToStr = Tag_NyStr_ObjAp("FmtWs", "Lvl Hid Wdt Bdr Cor Fmt   Lbl Tit    Bet Fmt Sum Avg Cnt", _
'    Lvl, Hid, Wdt, Bdr, Cor, Fmt, _
'    Lbl, Tit, _
'    Bet, Fml, Sum, Avg, Cnt)
'End Property
'
'Friend Sub SetBet(A As CnoVals)
'Set B_Bet = A
'End Sub
'
'Friend Sub SetTot(SumCny$(), AvgCny$(), CntCny$())
'B_SumCny = SumCny
'B_AvgCny = AvgCny
'B_CntCny = CntCny
'End Sub
'Friend Sub SetFmt(A As CnoVals)
'Set B_Fmt = A
'End Sub
'Friend Sub SetLvl(A As CnoVals)
'Set B_Lvl = A
'End Sub
'Friend Sub SetTit(A As CnoVals)
'Set B_Tit = A
'End Sub
'Friend Sub SetLbl(A As CnoVals)
'Set B_Lbl = A
'End Sub
'
'Friend Sub SetFml(A As CnoVals)
'Set B_Fml = A
'End Sub
'
'Sub SetHid(Cno%())
'B_HidCnoAy = Cno
'End Sub
'
'Friend Sub Tst()
'ZZ_DoFmt
'End Sub
'
'Private Sub ZZ_DoFmt()
'Dim A As New LoFmtr
'Dim B As LoFmtrRslt
'Set B = A.InitBySampleLy.Validate
'B.FmtWs.DoFmt SampleLo
''A.InitBySampleLy.Validate.FmtWs.DoFmt SampleLo
'End Sub
'
'Sub DoFmt(A As ListObject)
'DoFmtBdr A
'DoFmtFmt A
'DoFmtCor A
'DoFmtHid A
'DoFmtFml A
'DoFmtLvl A
'DoFmtBet A
'DoFmtTit A
'DoFmtTot A
'DoFmtLbl A
'End Sub
'
'Sub DoFmtBet(A As ListObject)
'Dim Ay() As CnoVal, C%, V, J%, Fml$
'Ay = B_Bet.Ay
'For J = 0 To UB(Ay)
'    C = Ay(J).Cno
'    With Brk(Ay(J).V, " ")
'        Fml = FmtQQ("=Sum([?]:[?])", .S1, .S2)
'    End With
'    LoC(A, C).Formula = Fml
'Next
'End Sub
'Sub DoFmtBdr(A As ListObject)
'Dim J%, C%(): C = B_BdrCnoAy
'For J = 0 To UB(C)
'    RgBdrLeft LoC(A, C(J))
'Next
'End Sub
'Sub DoFmtFmt(A As ListObject)
'Dim Ay() As CnoVal, C%, V, J%
'Ay = B_Fmt.Ay
'For J = 0 To UB(Ay)
'    C = Ay(J).Cno
'    V = Ay(J).V
'    LoC(A, C).NumberFormat = V
'Next
'End Sub
'Sub DoFmtCor(A As ListObject)
'Dim Ay() As CnoVal, C%, V, J%
'Ay = B_Cor.Ay
'For J = 0 To UB(Ay)
'    C = Ay(J).Cno
'    V = Ay(J).V
'    LoC(A, C).Interior.Color = V
'Next
'End Sub
'Sub DoFmtFml(A As ListObject)
'Dim Ay() As CnoVal, C%, V, J%
'Ay = B_Fml.Ay
'For J = 0 To UB(B_Fml)
'    C = Ay(J).Cno
'    V = Ay(J).V
'    LoC(A, C).Formula = V
'Next
'End Sub
'Sub DoFmtHid(A As ListObject)
'Dim J%, C%(): C = B_HidCnoAy
'For J = 0 To UB(C)
'    LoC(A, C(J)).EntireColumn.Hidden = True
'Next
'End Sub
'Sub DoFmtLvl(A As ListObject)
'LoWs(A).Outline.SummaryColumn = xlSummaryOnLeft
'Dim Ay() As CnoVal, C%, V, J%
'Ay = B_Lvl.Ay
'For J = 0 To UB(Ay)
'    C = Ay(J).Cno
'    V = Ay(J).V
'    LoC(A, C).EntireColumn.OutlineLevel = V
'Next
'End Sub
'
'Sub DoFmtTot(A As ListObject)
'Y A, B_SumCny, xlTotalsCalculationSum
'Y A, B_AvgCny, xlTotalsCalculationAverage
'Y A, B_CntCny, xlTotalsCalculationCount
'End Sub
'
'Sub DoFmtWdt(A As ListObject)
'Dim Ay() As CnoVal, C%, V, J%
'Ay = B_Wdt.Ay
'For J = 0 To UB(Ay)
'    C = Ay(J).Cno
'    V = Ay(J).V
'    LoC(A, C).ColumnWidth = V
'Next
'End Sub
'
'Private Function ZTitNRow%()
'Dim O%
'    Dim A$(), J%, M%
'    A$() = B_Tit.StrValAy
'    For J = 0 To UB(A)
'        M = Sz(Split(A(J), "|"))
'        If M > 0 Then O = M
'    Next
'ZTitNRow = O
'End Function
'
'Sub DoFmtTit(A As ListObject)
'Dim Fny$():         Fny = LoFny(A)
'Dim A1 As Range: Set A1 = A.DataBodyRange
'Dim HasTit As Boolean ' HasTit means if there is enough space above DtaRg to put the title
'    Dim N%: N = ZTitNRow
'    Dim TitR%: TitR = A1.Row - 2 - N ' Tit-Row
'    HasTit = TitR >= 0
'If Not HasTit Then Exit Sub
'
'Dim At As Range         ' At is TitAt
'    Set At = WsRC(RgWs(A), TitR, A1.Column) ' Use Ws to find this 'At', because TitR is relative to Ws
'    Dim Sq(): Sq = ZTitSq(Fny)
'TitRg_Fmt SqRg(Sq, At)   '<-- put the title and fmt the title Range
'End Sub
'
'Sub DoFmtLbl(A As ListObject)
'Dim Lbl$, Ay() As CnoVal, J%, C%
'Dim RFld As Range
'Dim RLbl As Range
'Ay = B_Lbl.Ay
'For J = 0 To UB(Ay)
'    C = Ay(J).Cno
'    Lbl = Ay(J).V
'    Set RFld = RgRC(LoC(A, C), 0, 1)
'    Set RLbl = RgRC(RFld, 0, 1)
'    RLbl.Value = RFld.Value    '<-- Swapping
'    RFld.Value = Lbl    '<-- Swapping
'Next
'End Sub
'Private Function ZTitDry(Fny$()) As Variant()
''From B_Tit & Fny, return TitDry
''If some column has no title, use FldNm as Tit
'Dim O(), J%, Ix%, Ay() As CnoVal
'For J = 0 To UB(Fny)
'    Ix = B_Tit.CnoIx(J)
'    If Ix = -1 Then
'        Push O, Array(Fny(J))
'    Else
'        Push O, SplitVBar(Ay(J).V, Trim:=True)  ' V contains Tit of current fld
'    End If
'Next
'ZTitDry = O
'End Function
'
'Private Function ZTitSq(Fny$()) As Variant()
'Dim TitSq(): TitSq = DrySq(ZTitDry(Fny))
'ZTitSq = SqTranspose(TitSq)
'End Function
'Private Sub Y(Lo As ListObject, C$(), A As XlTotalsCalculation)
'Dim J%
'For J = 0 To UB(C)
'    Lo.ListColumns(C(J)).TotalsCalculation = A
'Next
'End Sub

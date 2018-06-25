Attribute VB_Name = "M_FmtWs"
Option Explicit
Sub LoFmtr_FmtLo(A As WsFmtr, Lo As ListObject)

End Sub

Private Sub SetBet(A() As CnoVal)
Set B_Bet = A
End Sub

Private Sub SetTot(SumCny$(), AvgCny$(), CntCny$())
B_SumCny = SumCny
B_AvgCny = AvgCny
B_CntCny = CntCny
End Sub
Private Sub SetFmt(A() As CnoVal)
Set B_Fmt = A
End Sub
Private Sub SetLvl(A() As CnoVal)
Set B_Lvl = A
End Sub
Private Sub SetTit(A() As CnoVal)
Set B_Tit = A
End Sub
Private Sub SetLbl(A() As CnoVal)
Set B_Lbl = A
End Sub

Private Sub SetFml(A() As CnoVal)
Set B_Fml = A
End Sub

Sub SetHid(Cno%())
B_HidCnoAy = Cno
End Sub

Private Sub Tst()
ZZ_DoFmt
End Sub

Private Sub ZZ_DoFmt()
Dim A As New LoFmtr
Dim B As LoFmtrRslt
Set B = A.InitBySampleLy.Validate
B.FmtWs.DoFmt SampleLo
'A.InitBySampleLy.Validate.FmtWs.DoFmt SampleLo
End Sub

Sub DoFmt(A As ListObject)
DoFmtBdr A
DoFmtFmt A
DoFmtCor A
DoFmtHid A
DoFmtFml A
DoFmtLvl A
DoFmtBet A
DoFmtTit A
DoFmtTot A
DoFmtLbl A
End Sub

Sub DoFmtBet(A As ListObject)
Dim Ay() As CnoVal, C%, V, J%, Fml$
Ay = B_Bet.Ay
For J = 0 To UB(Ay)
    C = Ay(J).Cno
    With Brk(Ay(J).V, " ")
        Fml = FmtQQ("=Sum([?]:[?])", .S1, .S2)
    End With
    LoC(A, C).Formula = Fml
Next
End Sub
Sub DoFmtBdr(A As ListObject)
Dim J%, C%(): C = B_BdrCnoAy
For J = 0 To UB(C)
    RgBdrLeft LoC(A, C(J))
Next
End Sub
Sub DoFmtFmt(A As ListObject)
Dim Ay() As CnoVal, C%, V, J%
Ay = B_Fmt.Ay
For J = 0 To UB(Ay)
    C = Ay(J).Cno
    V = Ay(J).V
    LoC(A, C).NumberFormat = V
Next
End Sub
Sub DoFmtCor(A As ListObject)
Dim Ay() As CnoVal, C%, V, J%
Ay = B_Cor.Ay
For J = 0 To UB(Ay)
    C = Ay(J).Cno
    V = Ay(J).V
    LoC(A, C).Interior.Color = V
Next
End Sub
Sub DoFmtFml(A As ListObject)
Dim Ay() As CnoVal, C%, V, J%
Ay = B_Fml.Ay
For J = 0 To UB(B_Fml)
    C = Ay(J).Cno
    V = Ay(J).V
    LoC(A, C).Formula = V
Next
End Sub
Sub DoFmtHid(A As ListObject)
Dim J%, C%(): C = B_HidCnoAy
For J = 0 To UB(C)
    LoC(A, C(J)).EntireColumn.Hidden = True
Next
End Sub
Sub DoFmtLvl(A As ListObject)
LoWs(A).Outline.SummaryColumn = xlSummaryOnLeft
Dim Ay() As CnoVal, C%, V, J%
Ay = B_Lvl.Ay
For J = 0 To UB(Ay)
    C = Ay(J).Cno
    V = Ay(J).V
    LoC(A, C).EntireColumn.OutlineLevel = V
Next
End Sub

Sub DoFmtTot(A As ListObject)
Y A, B_SumCny, xlTotalsCalculationSum
Y A, B_AvgCny, xlTotalsCalculationAverage
Y A, B_CntCny, xlTotalsCalculationCount
End Sub

Sub DoFmtWdt(A As ListObject)
Dim Ay() As CnoVal, C%, V, J%
Ay = B_Wdt.Ay
For J = 0 To UB(Ay)
    C = Ay(J).Cno
    V = Ay(J).V
    LoC(A, C).ColumnWidth = V
Next
End Sub

Private Function ZTitNRow%()
Dim O%
    Dim A$(), J%, M%
    A$() = B_Tit.StrValAy
    For J = 0 To UB(A)
        M = Sz(Split(A(J), "|"))
        If M > 0 Then O = M
    Next
ZTitNRow = O
End Function

Sub DoFmtTit(A As ListObject)
Dim Fny$():         Fny = LoFny(A)
Dim A1 As Range: Set A1 = A.DataBodyRange
Dim HasTit As Boolean ' HasTit means if there is enough space above DtaRg to put the title
    Dim N%: N = ZTitNRow
    Dim TitR%: TitR = A1.Row - 2 - N ' Tit-Row
    HasTit = TitR >= 0
If Not HasTit Then Exit Sub

Dim At As Range         ' At is TitAt
    Set At = WsRC(RgWs(A), TitR, A1.Column) ' Use Ws to find this 'At', because TitR is relative to Ws
    Dim Sq(): Sq = ZTitSq(Fny)
TitRg_Fmt SqRg(Sq, At)   '<-- put the title and fmt the title Range
End Sub

Sub DoFmtLbl(A As ListObject)
Dim Lbl$, Ay() As CnoVal, J%, C%
Dim RFld As Range
Dim RLbl As Range
Ay = B_Lbl.Ay
For J = 0 To UB(Ay)
    C = Ay(J).Cno
    Lbl = Ay(J).V
    Set RFld = RgRC(LoC(A, C), 0, 1)
    Set RLbl = RgRC(RFld, 0, 1)
    RLbl.Value = RFld.Value    '<-- Swapping
    RFld.Value = Lbl    '<-- Swapping
Next
End Sub
Private Function ZTitDry(Fny$()) As Variant()
'From B_Tit & Fny, return TitDry
'If some column has no title, use FldNm as Tit
Dim O(), J%, Ix%, Ay() As CnoVal
For J = 0 To UB(Fny)
    Ix = B_Tit.CnoIx(J)
    If Ix = -1 Then
        Push O, Array(Fny(J))
    Else
        Push O, SplitVBar(Ay(J).V, Trim:=True)  ' V contains Tit of current fld
    End If
Next
ZTitDry = O
End Function

Private Function ZTitSq(Fny$()) As Variant()
Stop
'Dim TitSq(): TitSq = DrySq(ZTitDry(Fny))
'ZTitSq = SqTranspose(TitSq)
End Function
Private Sub Y(Lo As ListObject, C$(), A As XlTotalsCalculation)
Dim J%
For J = 0 To UB(C)
    Lo.ListColumns(C(J)).TotalsCalculation = A
Next
End Sub




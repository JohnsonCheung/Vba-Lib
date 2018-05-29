VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FmtWs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private B_BdrL%(), B_BdrR%()
Private B_BetC%(), B_BetA$(), B_BetB$()
Private B_CorC%(), B_Cor&()
Private B_FmlC%(), B_Fml$()
Private B_FmtC%(), B_Fmt$()
Private B_HidC%()
Private B_LblC%(), B_Lbl$()
Private B_LvlC%(), B_Lvl() As Byte
Private B_Sum$(), B_Avg$(), B_Cnt$() 'FldNm
Private B_WdtC%(), B_Wdt%()

Property Get ToStr$()
Dim O$()

End Property

Sub SetBdr(BdrL%(), BdrR%())
B_BdrL = BdrL
B_BdrR = BdrR
End Sub

Sub SetBet(BetC%(), BetA$(), BetB$())
B_BetC = BetC
B_BetA = BetA
B_BetB = BetB
End Sub

Sub SetTot(Sum$(), Avg$(), Cnt$())
B_Sum = Sum
B_Avg = Avg
B_Cnt = Cnt
End Sub
Sub SetFmt(FmtC%(), Fmt$())
B_FmtC = FmtC
B_Fmt = Fmt
End Sub
Sub SetLvl(LvlC%(), Lvl() As Byte)
B_LvlC = LvlC
B_Lvl = Lvl
End Sub
Sub SetTit(TitC%(), Tit$())
B_TitC = TitC
B_Tit = Tit
End Sub
Sub SetLbl(LblC%(), Lbl$())
B_LblC = LblC
B_Lbl = Lbl
End Sub
Sub SetFml(FmlC%(), Fml$())
B_FmlC = FmlC
B_Fml = Fml
End Sub
Sub SetHid(HidC%())
B_HidC = HidC
End Sub
Sub SetWdt(WdtC%(), Wdt%())
B_WdtC = WdtC
B_Wdt = Wdt
End Sub
Sub SetWdt1(A As CnoAyWdtAy)
B_WdtC = A.CnoAy
B_Wdt = A.WdtAy
End Sub
Sub SetCor(CorC%(), Cor&())
B_CorC = CorC
B_Cor = Cor
End Sub

Sub DoFmt(A As ListObject)
Dim DtaRg As Range
Dim J%
For J = 0 To UB(B_BdrL)
    RgBdrLeft RgC(DtaRg, B_BdrL(J))
Next
For J = 0 To UB(B_BdrR)
    RgBdrRight RgC(DtaRg, B_BdrR(J))
Next
For J = 0 To UB(B_Fmt)
    RgC(DtaRg, B_FmtC(J)).NumberFormat = B_Fmt(J)
Next
For J = 0 To UB(B_Cor)
    RgC(DtaRg, B_CorC(J)).Interior.Color = B_Cor(J)
Next
For J = 0 To UB(B_Fml)
    RgC(DtaRg, B_Fml(J)).Formula = B_Fml(J)
Next
For J = 0 To UB(B_HidC)
    RgC(DtaRg, B_HidC(J)).Hidden = True
Next
For J = 0 To UB(B_Lvl)
    RgC(DtaRg, B_LvlC(J)).EntireColumn.OutlineLevel = B_Lvl(J)
Next
Y A, B_Sum, xlTotalsCalculationSum
Y A, B_Avg, xlTotalsCalculationAverage
Y A, B_Cnt, xlTotalsCalculationCount
For J = 0 To UB(B_Cno)
    RgC(DtaRg, B_Cno(J)).ColumnWidth = B_Wdt(J)
Next
End Sub
Private Sub Y(Lo As ListObject, C$(), A As XlTotalsCalculation)
Dim J%
For J = 0 To UB(C)
    Lo.ListColumns(C(J)).TotalsCalculation = A
Next
End Sub

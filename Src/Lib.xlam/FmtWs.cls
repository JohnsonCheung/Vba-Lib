VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FmtWs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private B_WdtC%()
Private B_Wdt%()
Private B_Sum$() 'FldNm
Private B_Avg$()
Private B_Cnt$()
Private B_FmtC%()
Private B_Fmt$()
Private B_LvlC%()
Private B_Lvl() As Byte
Private B_LblC%()
Private B_Lbl$()
Private B_HidC%()
Private B_FmlC%()
Private B_Fml$()
Private B_CorC%()
Private B_Cor&()
Private B_BetC%()
Private B_BetA$()
Private B_BetB$()
Private B_BdrL%()
Private B_BdrR%()
Property Get ToStr$()

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
Sub SetCor(CorC%(), Cor&())
B_CorC = CorC
B_Cor = Cor
End Sub

Sub DoFmt(A As ListObject)
Dim DtaRg As Range
Dim J%
For J = 0 To UB(B_Cno)
    RgC(DtaRg, B_Cno(J)).NumberFormat = B_Fmt(J)
Next
For J = 0 To UB(B_Cno)
    RgC(DtaRg, B_Cno(J)).Interior.Color = B_Cor(J)
Next
For J = 0 To UB(B_Cno)
    RgC(DtaRg, B_Cno(J)).Formula = B_Fml(J)
Next
For J = 0 To UB(B_Cno)
    RgC(DtaRg, B_Cno(J)).NumberFormat = B_Fmt(J)
Next
For J = 0 To UB(B_Cno)
    RgC(DtaRg, B_Cno(J)).Hidden = True
Next
For J = 0 To UB(B_LvlC)
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

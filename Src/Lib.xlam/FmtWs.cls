VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FmtWs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private A_WdtC%()
Private A_Wdt%()
Private A_Sum$() 'FldNm
Private A_Avg$()
Private A_Cnt$()
Private A_FmtC%()
Private A_Fmt$()
Private A_LvlC%()
Private A_Lvl() As Byte
Private A_LblC%()
Private A_Lbl$()
Private A_HidC%()
Private A_FmlC%()
Private A_Fml$()
Private A_CorC%()
Private A_Cor&()
Private A_BetC%()
Private A_BetA$()
Private A_BetB$()
Private A_BdrL%()
Private A_BdrR%()
Function Show() As String()

End Function
Sub SetBdr(BdrL%(), BdrR%())
A_BdrL = BdrL
A_BdrR = BdrR
End Sub
Sub SetBet(BetC%(), BetA$(), BetB$())
A_BetC = BetC
A_BetA = BetA
A_BetB = BetB
End Sub
Sub SetTot(Sum$(), Avg$(), Cnt$())
A_Sum = Sum
A_Avg = Avg
A_Cnt = Cnt
End Sub
Sub SetFmt(FmtC%(), Fmt$())
A_FmtC = FmtC
A_Fmt = Fmt
End Sub
Sub SetLvl(LvlC%(), Lvl() As Byte)
A_LvlC = LvlC
A_Lvl = Lvl
End Sub
Sub SetTit(TitC%(), Tit$())
A_TitC = TitC
A_Tit = Tit
End Sub
Sub SetLbl(LblC%(), Lbl$())
A_LblC = LblC
A_Lbl = Lbl
End Sub
Sub SetFml(FmlC%(), Fml$())
A_FmlC = FmlC
A_Fml = Fml
End Sub
Sub SetHid(HidC%())
A_HidC = HidC
End Sub
Sub SetWdt(WdtC%(), Wdt%())
A_WdtC = WdtC
A_Wdt = Wdt
End Sub
Sub SetCor(CorC%(), Cor&())
A_CorC = CorC
A_Cor = Cor
End Sub

Sub DoFmt(A As ListObject)
Dim DtaRg As Range
Dim J%
For J = 0 To UB(A_Cno)
    RgC(DtaRg, A_Cno(J)).NumberFormat = A_Fmt(J)
Next
For J = 0 To UB(A_Cno)
    RgC(DtaRg, A_Cno(J)).Interior.Color = A_Cor(J)
Next
For J = 0 To UB(A_Cno)
    RgC(DtaRg, A_Cno(J)).Formula = A_Fml(J)
Next
For J = 0 To UB(A_Cno)
    RgC(DtaRg, A_Cno(J)).NumberFormat = A_Fmt(J)
Next
For J = 0 To UB(A_Cno)
    RgC(DtaRg, A_Cno(J)).Hidden = True
Next
For J = 0 To UB(A_LvlC)
    RgC(DtaRg, A_LvlC(J)).EntireColumn.OutlineLevel = A_Lvl(J)
Next
Y A, A_Sum, xlTotalsCalculationSum
Y A, A_Avg, xlTotalsCalculationAverage
Y A, A_Cnt, xlTotalsCalculationCount
For J = 0 To UB(A_Cno)
    RgC(DtaRg, A_Cno(J)).ColumnWidth = A_Wdt(J)
Next
End Sub
Private Sub Y(Lo As ListObject, C$(), A As XlTotalsCalculation)
Dim J%
For J = 0 To UB(C)
    Lo.ListColumns(C(J)).TotalsCalculation = A
Next
End Sub

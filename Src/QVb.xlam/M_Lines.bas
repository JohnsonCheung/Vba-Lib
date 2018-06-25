Attribute VB_Name = "M_Lines"
Option Explicit
Public Lines$

Property Get LinesEndTrim$(Lines)
LinesEndTrim = JnCrLf(LyEndTrim(SplitCrLf(Lines)))
End Property

Property Get LinesLasLin$(Lines)
Dim A$(): A = LinesLy(Lines): If Sz(A) = 0 Then Exit Property
LinesLasLin = AyLasEle(A)
End Property

Property Get LinesLinCnt&(Lines)
LinesLinCnt = Sz(SplitCrLf(Lines))
End Property

Property Get LinesLy(Lines) As String()
LinesLy = SplitLines(Lines)
End Property

Property Get LinesSqH(Lines) As Variant()
LinesSqH = AySqH(LinesLy(Lines))
End Property

Property Get LinesSqV(Lines) As Variant()
LinesSqV = AySqV(LinesLy(Lines))
End Property

Property Get LinesVbl$(Lines)
If InStr(Lines, "|") Then Er "Lines.ToVbl", "Cannt have [|] in {Lines}", Lines
LinesVbl = Replace(Lines, vbCrLf, "|")
End Property

Property Get LinesWdt%(Lines)
LinesWdt = AyWdt(SplitCrLf(Lines))
End Property

Sub ZZ__Tst()
ZZ_LinesEndTrim
End Sub

Private Sub ZZ_LinesEndTrim()
Dim Lines$: Lines = RplVBar("lksdf|lsdfj|||")
Dim Act$: Act = LinesEndTrim(Lines)
Debug.Print Act & "<"
Stop
End Sub

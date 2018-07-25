Attribute VB_Name = "M_Lines"
Option Explicit
Public Lines$

Function LinesEndTrim$(Lines)
LinesEndTrim = JnCrLf(LyEndTrim(SplitCrLf(Lines)))
End Function

Function LinesLasLin$(Lines)
Dim A$(): A = LinesLy(Lines): If Sz(A) = 0 Then Exit Function
LinesLasLin = AyLasEle(A)
End Function

Function LinesLinCnt&(Lines)
LinesLinCnt = Sz(SplitCrLf(Lines))
End Function

Function LinesLy(Lines) As String()
LinesLy = SplitLines(Lines)
End Function

Function LinesSqH(Lines) As Variant()
LinesSqH = AySqH(LinesLy(Lines))
End Function

Function LinesSqV(Lines) As Variant()
LinesSqV = AySqV(LinesLy(Lines))
End Function

Function LinesVbl$(Lines)
If InStr(Lines, "|") Then Er "Lines.ToVbl", "Cannt have [|] in {Lines}", Lines
LinesVbl = Replace(Lines, vbCrLf, "|")
End Function

Function LinesWdt%(Lines)
LinesWdt = AyWdt(SplitCrLf(Lines))
End Function

Sub ZZZ__Tst()
ZZ_LinesEndTrim
End Sub

Private Sub ZZ_LinesEndTrim()
Dim Lines$: Lines = RplVBar("lksdf|lsdfj|||")
Dim Act$: Act = LinesEndTrim(Lines)
Debug.Print Act & "<"
Stop
End Sub

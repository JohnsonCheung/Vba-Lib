VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Lines$

Property Get LasLin$()
Dim A$(): A = Ly: If Sz(A) = 0 Then Exit Property
LasLin = AyLasEle(A)
End Property

Property Get LinCnt&()
LinCnt = Sz(SplitCrLf(Lines))
End Property

Property Get Ly() As String()
Ly = SplitLines(Lines)
End Property

Property Get Vbl$()
If InStr(Lines, "|") Then Er "Lines.ToVbl", "Cannt have [|] in {Lines}", Lines
Vbl = Replace(Lines, vbCrLf, "|")
End Property

Property Get Wdt%()
Wdt = AyWdt(Ly)
End Property

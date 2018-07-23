Attribute VB_Name = "M_Fny"
Option Explicit

Sub FnyWhFldLvs(Fny$(), FldLvs$, ParamArray OAp())
'FnyWhFldLvs=Field Index Array
Dim A$(): A = SplitSpc(FldLvs)
Dim I&(): I = AyIxAy(Fny, A)
Dim J%
For J = 0 To UB(I)
    OAp(J) = I(J)
Next
End Sub

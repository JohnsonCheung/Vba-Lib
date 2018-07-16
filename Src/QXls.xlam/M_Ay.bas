Attribute VB_Name = "M_Ay"
Option Explicit

Function AyRgH(Ay, At As Range) As Range
Set AyRgH = SqRg(AySqH(Ay), At)
End Function

Function AyRgV(Ay, At As Range) As Range
Set AyRgV = SqRg(AySqV(Ay), At)
End Function


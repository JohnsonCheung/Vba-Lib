Attribute VB_Name = "M_Ap"
Option Explicit

Function ApIntAy(ParamArray Ap()) As Integer()
Dim Av(): Av = Ap
ApIntAy = AyIntAy(Av)
End Function

Function ApLngAy(ParamArray Ap()) As Long()
Dim Av(): Av = Ap
ApLngAy = AyLngAy(Av)
End Function

Function ApSngAy(ParamArray Ap()) As Single()
Dim Av(): Av = Ap
ApSngAy = AySngAy(Av)
End Function

Function ApSy(ParamArray Ap()) As String()
Dim Av(): Av = Ap
ApSy = AySy(Av)
End Function

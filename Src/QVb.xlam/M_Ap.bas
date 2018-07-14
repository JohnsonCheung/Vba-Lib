Attribute VB_Name = "M_Ap"
Option Explicit

Property Get ApIntAy(ParamArray Ap()) As Integer()
Dim Av(): Av = Ap
ApIntAy = AyIntAy(Av)
End Property

Property Get ApLngAy(ParamArray Ap()) As Long()
Dim Av(): Av = Ap
ApLngAy = AyLngAy(Av)
End Property

Property Get ApSngAy(ParamArray Ap()) As Single()
Dim Av(): Av = Ap
ApSngAy = AySngAy(Av)
End Property

Property Get ApSy(ParamArray Ap()) As String()
Dim Av(): Av = Ap
ApSy = AySy(Av)
End Property


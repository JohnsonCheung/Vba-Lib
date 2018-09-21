Attribute VB_Name = "M_Ap"
Option Explicit
Function ApDtAy(ParamArray Ap()) As Dt()
Dim Av(): Av = Ap
ApDtAy = AyInto(Av, EmpDtAy)
End Function
Private Sub ZZ_ApDtAy()
Dim A() As Dt
A = ApDtAy(SampleDt1, SampleDt2)
Stop
End Sub

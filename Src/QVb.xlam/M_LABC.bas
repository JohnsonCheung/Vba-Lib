Attribute VB_Name = "M_LABC"
Option Explicit

Property Get Seed_Expand$(QVbl$, Ny0)
Dim O$()
Dim Sy$()
    Sy = SplitVBar(QVbl)
Dim Ny$(), J%, I%
Ny = DftNy(Ny0)
For J = 0 To UB(Ny)
    For I = 0 To UB(Sy)
       Push O, Replace(Sy(I), "?", Ny(J))
    Next
Next
Seed_Expand = JnCrLf(O)
End Property

Private Sub ZZ_Seed_Expand()
Dim Ny0
Dim QVbl$
QVbl = "Sub Tst?()|Dim A As New ?: A.Tst|End Sub"
Ny0 = "Xws Xwb Xfx Xrg"
Debug.Print Seed_Expand(QVbl, Ny0)
End Sub

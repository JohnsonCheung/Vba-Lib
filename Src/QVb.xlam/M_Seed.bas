Attribute VB_Name = "M_Seed"
Option Explicit

Property Get Seed_Expand$(VblQQStr, Ny0)
'Seed is a VblQQ-String
Dim A$, J%, O$()
Dim Ny$()
Ny = DftNy(Ny0)
For J = 0 To UB(Ny)
    Push O, Replace(VblQQStr, "?", Ny(J))
Next
Seed_Expand = RplVBar(JnCrLf(O))
End Property

Private Sub ZZ_Seed_Expand()
Dim Ny0
Dim QVbl$
QVbl = "Sub Tst?()|Dim A As New ?: A.Tst|End Sub"
Ny0 = "Xws Xwb Xfx Xrg"
Debug.Print Seed_Expand(QVbl, Ny0)
End Sub

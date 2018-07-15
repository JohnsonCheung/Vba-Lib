VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Seed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private B_SeedSy$()

Function Expand$(Ny0)
Dim O$()
Dim Ny$(), J%, I%
Ny = DftNy(Ny0)
For J = 0 To UB(Ny)
    For I = 0 To UB(B_SeedSy)
       Push O, Replace(B_SeedSy(I), "?", Ny(J))
    Next
Next
Expand = JnCrLf(O)
End Function

Friend Property Get Init(Seed0) As Seed
B_SeedSy = DftLy(Seed0)
Set Init = Me
End Property

Friend Sub Tst()
Expand__Tst
End Sub

Private Sub Expand__Tst()
Dim Tp$
Dim Seed0$
Tp = "Sub Tst?()|Dim A As New ?: A.Tst|End Sub"
Seed0 = "Xws Xwb Xfx Xrg"
Debug.Print Vb.Seed(Seed0).Expand(Tp)
End Sub

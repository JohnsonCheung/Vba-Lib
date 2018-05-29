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
Friend Function Init(Seed0) As Seed
B_SeedSy = DftNy(Seed0)
Set Init = Me
End Function
Function Expand$(Tp0)
Dim TpLy$(): TpLy = DftTpLy(Tp0)
Expand = Expand1(TpLy)
End Function

Private Function Expand1$(Ly$())
Dim O$(), J%, I%
Stop
For I = 0 To UB(B_SeedSy)
   For J = 0 To UB(Ly)
       Push O, Replace(Ly(J), "?", B_SeedSy(I))
   Next
Next
Expand1 = JnCrLf(O)
End Function
Friend Sub Tst()
Expand1__Tst
End Sub
Private Sub Expand1__Tst()
Dim Tp$
Dim Seed0$
Tp = "Sub Tst?()|Dim A As New ?: A.Tst|End Sub"
Seed0 = "Xws Xwb Xfx Xrg"
Debug.Print Vb.Seed(Seed0).Expand(Tp)
End Sub



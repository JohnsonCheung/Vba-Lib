Attribute VB_Name = "VbSeed"
Option Explicit

Function SeedExp$(Seed0, Tp0)
Dim Seed$(): Seed = DftNy(Seed0)
Dim Tp$()
Select Case True
Case VarIsStr(Tp0): Tp = SplitVBar(Tp0)
Case VarIsSy(Tp0):  Tp = Tp0
Case Else: Stop
End Select
SeedExp = SeedExp__Expand(Seed, Tp)
End Function

Private Function SeedExp__Expand$(Seed$(), Ly$())
Dim O$(), J%, I%
For I = 0 To UB(Seed)
   For J = 0 To UB(Ly)
       Push O, Replace(Ly(J), "?", Seed(I))
   Next
Next
SeedExp__Expand = JnCrLf(O)
End Function

Private Sub SeedExpLvs__Tst()
Dim Tp$
Dim Seed$
Tp = "Sub Tst?()|Dim A As New ?: A.Tst|End Sub"
Seed = "Xws Xwb Xfx Xrg"
Debug.Print SeedExp(Seed, Tp)
End Sub

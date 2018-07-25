Attribute VB_Name = "F_Ide_Z3DDMth"
Option Explicit
Sub Gen_Md_Z3DDTst()
Dim M As CodeModule: Set M = CurMd
MdGenZ3DDTst M
MdEnsZ3DMthAsPrivate M
End Sub
Function MdGenZ3DDTst(A As CodeModule)
MthRmv Mth(A, "ZZZ__Tst")
MdAppLines A, MdZ3DDTstMthLines(A)
End Function

Function MdZ3DDTstMthLines$(A As CodeModule)
Dim Ay$(): Ay = MdMthNy(A, "^ZZZ_", IsNoMdNmPfx:=True)
Dim Ay1$(): Ay1 = AyRmvEle(Ay, "ZZZ__Tst")
Dim Ay2$(): Ay2 = AySrt(Ay1)
Dim O$(), J%
Push O, "Sub ZZZ__Tst()"
For J = 0 To UB(Ay2)
    Push O, Ay2(J)
Next
Push O, "End Sub"
MdZ3DDTstMthLines = JnCrLf(O)
End Function
Private Sub ZZZ__A()

End Sub
Private Sub ZZZ__B()

End Sub
Sub ZZZ__Tst()
ZZZ__A
ZZZ__B
End Sub

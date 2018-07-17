Attribute VB_Name = "F_EnsPrivateTstMth"
Option Explicit

Sub MdEnsTstMthAsPrivate(A As CodeModule)
Dim Ay() As S1S2, S1S2 As S1S2
Dim MdNm$, MthNm$
Ay = MdNonPrivateTstMthS1S2Ay(A)
Dim J%, M As CodeModule
For J = 0 To S1S2_UB(Ay)
   S1S2 = Ay(J)
   MdNm = S1S2.S1
   MthNm = S1S2.S2
'   Set M = Md(MdNm)
   MthEnsPrivate M, MthNm
Next
End Sub

Sub MthEnsPrivate(A As CodeModule, MthNm$)

End Sub

Private Function MdNonPrivateTstMthS1S2Ay(A As CodeModule) As S1S2()

End Function

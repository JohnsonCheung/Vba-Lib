Attribute VB_Name = "SalRpt__SqlOf_TCrd__Tst"
Option Explicit
Private Type ZZ3_TstDta
   LisCrd As String
   BrkCrd As Boolean
   Exp As String
End Type

Private Function ZZ3_TstDta1() As ZZ3_TstDta
With ZZ3_TstDta1
   .BrkCrd = False
   .LisCrd = "1 2"
   .Exp = ""
End With
End Function

Private Function ZZ3_TstDta2() As ZZ3_TstDta
With ZZ3_TstDta2
   .BrkCrd = True
   .LisCrd = "1 2"
   .Exp = "Select|    CrdTyId      Crd  ,|    CrdTyNm      CrdNm|  Into #Crd|  From JR_FrqMbrLis_#CrdTy()|  Where CrdTyId in (1,2)"
End With
End Function

Private Function ZZ3_TstDta3() As ZZ3_TstDta
With ZZ3_TstDta3
   .BrkCrd = True
   .LisCrd = ""
   .Exp = "Select|    CrdTyId      Crd  ,|    CrdTyNm      CrdNm|  Into #Crd|  From JR_FrqMbrLis_#CrdTy()"
End With
End Function

Private Sub ZZ3_Tstr(A As ZZ3_TstDta)
With A
   Dim Act$
   Act = SqLoFmtr_TCrd(.BrkCrd, .LisCrd)
   Ass IsEq(Act, .Exp)
End With
End Sub

Private Sub SqLoFmtr_TCrd__Tst()
ZZ3_Tstr ZZ3_TstDta1
ZZ3_Tstr ZZ3_TstDta2
ZZ3_Tstr ZZ3_TstDta3
End Sub

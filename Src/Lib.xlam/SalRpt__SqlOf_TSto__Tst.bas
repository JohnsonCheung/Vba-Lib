Attribute VB_Name = "SalRpt__SqlOf_TSto__Tst"
Option Explicit
Private Type ZZ5_TstDta
   LisSto As String
   BrkSto As Boolean
   Exp As String
End Type

Private Function ZZ5_TstDta1() As ZZ5_TstDta
With ZZ5_TstDta1
   .BrkSto = False
   .LisSto = "001 002"
   .Exp = ""
End With
End Function

Private Function ZZ5_TstDta2() As ZZ5_TstDta
With ZZ5_TstDta2
   .BrkSto = True
   .LisSto = "001 002"
   .Exp = "Select|    '0'+Loc_Code      Sto   ,|    Loc_Name          StoNm ,|    Loc_CName         StoCNm|  Into #Sto|  From Location|  Where '0'+Loc_Code in ('001','002')"
End With
End Function

Private Function ZZ5_TstDta3() As ZZ5_TstDta
With ZZ5_TstDta3
   .BrkSto = True
   .LisSto = ""
   .Exp = "Select|    '0'+Loc_Code      Sto   ,|    Loc_Name          StoNm ,|    Loc_CName         StoCNm|  Into #Sto|  From Location"
End With
End Function

Private Sub ZZ5_Tstr(A As ZZ5_TstDta)
Ass IsEq(Srp_TSto(A.BrkSto, A.LisSto), A.Exp)
End Sub

Private Sub Srp_TSto__Tst()
ZZ5_Tstr ZZ5_TstDta1
ZZ5_Tstr ZZ5_TstDta2
ZZ5_Tstr ZZ5_TstDta3
End Sub

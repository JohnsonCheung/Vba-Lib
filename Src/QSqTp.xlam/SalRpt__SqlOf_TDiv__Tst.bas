Attribute VB_Name = "SalRpt__SqlOf_TDiv__Tst"
Option Explicit
Private Type ZZ4_TstDta
   BrkDiv As Boolean
   LisDiv As String
   Exp As String
End Type

Private Sub Srp_TDiv__Tst()
ZZ4_Tstr ZZ4_TstDta1
ZZ4_Tstr ZZ4_TstDta2
ZZ4_Tstr ZZ4_TstDta3
End Sub

Private Function ZZ4_TstDta1() As ZZ4_TstDta
With ZZ4_TstDta1
   .BrkDiv = False
   .LisDiv = "01 02"
End With
End Function

Private Function ZZ4_TstDta2() As ZZ4_TstDta
With ZZ4_TstDta2
   .BrkDiv = True
   .LisDiv = "01 02"
   .Exp = "Select|    Dept + Division      Div   ,|    DivNm                DivNm ,|    Seq                  DivSeq,|    Status               DivSts|  Into #Div|  From Division|  Where Dept + Division in ('01','02')"
End With
End Function

Private Function ZZ4_TstDta3() As ZZ4_TstDta
With ZZ4_TstDta3
   .BrkDiv = True
   .LisDiv = ""
   .Exp = "Select|    Dept + Division      Div   ,|    DivNm                DivNm ,|    Seq                  DivSeq,|    Status               DivSts|  Into #Div|  From Division"
End With
End Function

Private Sub ZZ4_Tstr(A As ZZ4_TstDta)
With A
   Ass IsEq(Srp_TDiv(.BrkDiv, .LisDiv), .Exp)
End With
End Sub

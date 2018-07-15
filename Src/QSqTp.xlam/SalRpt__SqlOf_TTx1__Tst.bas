Attribute VB_Name = "SalRpt__SqlOf_TTx1__Tst"
Option Explicit
'Option Explicit
'Type ZZ7_TstDta
'   P As SrPm
'   Exp As String
'End Type
'
'Private Function ZZ7_Ay() As ZZ7_TstDta()
'Dim O() As ZZ7_TstDta
'ZZ7_Push O, ZZ7_TstDta1
'ZZ7_Push O, ZZ7_TstDta2
'ZZ7_Ay = O
'End Function
'
'Private Sub ZZ7_Push(O() As ZZ7_TstDta, M As ZZ7_TstDta)
'Dim N%: N = ZZ7_Sz(O)
'ReDim Preserve O(N)
'O(N) = M
'End Sub
'
'Private Function ZZ7_Sz%(A() As ZZ7_TstDta)
'On Error Resume Next
'ZZ7_Sz = UBound(A) + 1
'End Function
'
'Private Function ZZ7_TstDta1() As ZZ7_TstDta
'Dim O As ZZ7_TstDta
'With O
'   With .P
'       .BrkCrd = True
'       .BrkDiv = True
'       .BrkMbr = True
'       .BrkSto = True
'       .LisCrd = "1 2 3"
'       .LisDiv = "01 02 03"
'       .LisSto = "001 002 004"
'       .FmDte = "20170101"
'       .ToDte = "20170131"
'       .SumLvl = "D"
'   End With
'   .Exp = _
'       "Select|    Case When|      SHMCode Like '134234%' OR|      SHMCode Like '12323%'  THEN 1|      Else Case When|      SHMCode Like '2444%'    OR|      SHMCode Like '2443434%' OR|      SHMCode Like '24424%'   THEN 2|      Else Case When|      SHMCode Like '3%' THEN 3|      Else 4|      End End End                                                              Crd  ,|    Sum(SHAmount)                                                            Amt  ,|    Sum(SHQty)                                                               Qty  ,|    Count(SHInvoice + SHSDate + SHRef)                                       Cnt  ,|    Mbr-Expr                                                                 Mbr  ,|    Div-Expr                                                                 Div  ,|    Sto-Expr                                                                 Sto  ,|    SUBSTR(SHSDate,1,4)                                                      TxY  ,|    SUBSTR(SHSDate,5,2)" & _
'       "TxM  ,|    TxW-Expr                                                                 TxW  ,|    SUBSTR(SHSDate,7,2)                                                      TxD  ,|    SUBSTR(SHSDate,1,4)+'/'+SUBSTR(SHSDate,5,2)+'/'+SUBSTR(SHSDate,7,2)      TxDte|  Into #Tx|  From SaleHistory|  Where SHDate Between '20170101' and '20170131'|    And Case When|SHMCode Like '134234%' OR|SHMCode Like '12323%'  THEN 1|Else Case When|SHMCode Like '2444%'    OR|SHMCode Like '2443434%' OR|SHMCode Like '24424%'   THEN 2|Else Case When|SHMCode Like '3%' THEN 3|Else 4|End End End  in (1,2,3)|    And Div-Expr in ('01','02','03')|    And Sto-Expr in ('001','002','004')|  Group By|Case When|SHMCode Like '134234%' OR|SHMCode Like '12323%'  THEN 1|Else Case When|SHMCode Like '2444%'    OR|SHMCode Like '2443434%' OR|SHMCode Like '24424%'   THEN 2|Else Case When|SHMCode Like '3%' THEN 3|Else 4|End End End                                                         ,|Mbr-Expr" & _
'       ",|Div-Expr                                                            ,|Sto-Expr                                                            ,|SUBSTR(SHSDate,1,4)                                                 ,|SUBSTR(SHSDate,5,2)                                                 ,|SUBSTR(SHSDate,7,2)                                                 ,|SUBSTR(SHSDate,1,4)+'/'+SUBSTR(SHSDate,5,2)+'/'+SUBSTR(SHSDate,7,2)"
'End With
'ZZ7_TstDta1 = O
'End Function
'
'Private Function ZZ7_TstDta2() As ZZ7_TstDta
'Dim O As ZZ7_TstDta
'With O.P
'    .BrkCrd = True
'    .BrkDiv = True
'    .BrkMbr = True
'    .BrkSto = True
'    .LisCrd = "1 2 3"
'    .LisDiv = "01 02 03"
'    .LisSto = "001 002 004"
'    .FmDte = "20170101"
'    .ToDte = "20170131"
'    .SumLvl = "D"
'End With
'O.Exp = ""
'ZZ7_TstDta2 = O
'End Function
'
'Private Sub ZZ7_Tstr(A As ZZ7_TstDta)
'Dim ECrd$
'With A
'   Ass IsEq(Srp_TTx(A.P), .Exp)
'End With
'End Sub
'
'Private Function ZZ7_UB%(A() As ZZ7_TstDta)
'ZZ7_UB = ZZ7_Sz(A) - 1
'End Function
'
'Private Sub ZZ7__Srp_TTx__Tst()
'Dim Ay() As ZZ7_TstDta: Ay = ZZ7_Ay
'Dim J%
'For J = 0 To UBound(Ay)
'   ZZ7_Tstr Ay(J)
'Next
'End Sub

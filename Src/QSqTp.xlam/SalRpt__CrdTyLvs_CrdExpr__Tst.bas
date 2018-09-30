Attribute VB_Name = "SalRpt__CrdTyLvs_CrdExpr__Tst"
'Option Explicit
'Private Type ZZ1_CrdExpr_TstDta
'   CrdTyLvs As String
'   CrdPfxTyDry() As Variant
'   ShouldThow As Boolean
'   Exp As String
'End Type
'
'Private Function ZZ1_Act(A As ZZ1_CrdExpr_TstDta) As String
'With A
'   ZZ1_Act = CrdTyLvs_CrdExpr(.CrdTyLvs, .CrdPfxTyDry)
'End With
'End Function
'
'Private Function ZZ1_ActOpt(A As ZZ1_CrdExpr_TstDta) As SomStr
'On Error GoTo X
'ZZ1_ActOpt = SomStr(ZZ1_Act(A))
'Exit Function
'X:
'End Function
'
'Private Sub ZZ1_Push(O() As ZZ1_CrdExpr_TstDta, M As ZZ1_CrdExpr_TstDta)
'Dim N&: N = ZZ1_Sz(O)
'ReDim Preserve O(N)
'O(N) = M
'End Sub
'
'Private Function ZZ1_Sz&(A() As ZZ1_CrdExpr_TstDta)
'On Error Resume Next
'ZZ1_Sz = UBound(A) + 1
'End Function
'
'Private Function ZZ1_TstDta0() As ZZ1_CrdExpr_TstDta
'With ZZ1_TstDta0
'   .CrdPfxTyDry = ZZCrdPfxTyDry
'   .CrdTyLvs = "1 2 3"
'   .ShouldThow = False
'   .Exp = "Case When|SHMCode Like '134234%' OR|SHMCode Like '12323%'  THEN 1|Else Case When|SHMCode Like '2444%'    OR|SHMCode Like '2443434%' OR|SHMCode Like '24424%'   THEN 2|Else Case When|SHMCode Like '3%' THEN 3|Else 4|End End End "
'End With
'End Function
'
'Private Function ZZ1_TstDta1() As ZZ1_CrdExpr_TstDta
'With ZZ1_TstDta1
'   .CrdPfxTyDry = ZZCrdPfxTyDry
'   .CrdTyLvs = "1"
'   .ShouldThow = False
'   .Exp = ""
'End With
'End Function
'
'Private Function ZZ1_TstDta2() As ZZ1_CrdExpr_TstDta
'With ZZ1_TstDta2
'   .CrdPfxTyDry = ZZCrdPfxTyDry
'   .CrdTyLvs = ""
'   .ShouldThow = True
'   .Exp = ""
'End With
'End Function
'
'Private Function ZZ1_TstDta3() As ZZ1_CrdExpr_TstDta
'With ZZ1_TstDta3
'   .CrdPfxTyDry = ZZCrdPfxTyDry
'   .CrdTyLvs = ""
'   .ShouldThow = False
'   .Exp = ""
'End With
'End Function
'
'Private Function ZZ1_TstDtaAy() As ZZ1_CrdExpr_TstDta()
'Dim O() As ZZ1_CrdExpr_TstDta
'ZZ1_Push O, ZZ1_TstDta0
'ZZ1_Push O, ZZ1_TstDta1
'ZZ1_Push O, ZZ1_TstDta2
'ZZ1_Push O, ZZ1_TstDta3
'ZZ1_TstDtaAy = O
'End Function
'
'Private Sub ZZ1_Tstr(A As ZZ1_CrdExpr_TstDta)
'Dim M As SomStr
'M = ZZ1_ActOpt(A)
'With A
'   If .ShouldThow Then
'       If M.Som Then Stop
'   Else
'       If Not M.Som Then Stop
'       Ass .Exp = M.Str
'   End If
'End With
'End Sub
'
'Private Function ZZCrdPfxTyDry() As Variant()
'ZZCrdPfxTyDry = DryOf_CrdPfxTy
'End Function
'
'Private Function ZZCrdTyLvs$()
'ZZCrdTyLvs = "1 2 3"
'End Function
'
'Private Sub ZZ1__CrdExpr__Tst()
'Dim Ay() As ZZ1_CrdExpr_TstDta
'   Ay = ZZ1_TstDtaAy
'Dim J%
'For J = 0 To UBound(Ay)
'   If J = J Then
'       ZZ1_Tstr Ay(J)
'   End If
'Next
'End Sub

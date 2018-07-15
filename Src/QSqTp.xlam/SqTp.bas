Attribute VB_Name = "SqTp"
Option Explicit
'Option Explicit
'Const C_BlkTyLvs$ = "ER PM SW SQ RM"
'Const Msg_Sq_1_NotInEDic = "These items not found in ExprDic [?]"
'Const Msg_Sq_1_MustBe1or0 = "For %?xxx, 2nd term must be 1 or 0"
'Const QItmLvs0$ = "sel selDis gp"
'Const C_LvsOf_1$ = "into fm left jn" ' QItm = Sql-Phrase-Itm
'Const C_LvsOf_2$ = "upd set"
'Const C_LvsOf_3$ = "whExpr whBetStr whBetNbr whInStrLis whInNbrLis"
'Const QItmLvs4$ = "andExpr andBetStr andBetNbr andInStrLis andInNbrLis"
''Const QItmLvs$ = QItmLvs0 & QItmLvs1 & QItmLvs2 & QItmLvs3 & QItmLvs4
'Const OptionalQItm$ = "sel selDis wh* and*"
'Type Sw
'    Tbl As Dictionary
'    Fld As Dictionary
'End Type
'Type SwRslt
'    Sw As Sw
'    Er As TpEr
'End Type
'Type StmtRslt
'    Er As TpEr
'    Stmt As String
'End Type
'Type SqTpRslt
'    MsgLy() As String
'    Stmts As String
'End Type
'Private Type StmtsRslt
'    Er As TpEr
'    Stmts As String
'End Type
'Private Type PmRslt
'    Pm As Dictionary
'    Er As TpEr
'End Type
''============
'Private Type Blk
'    BlkTyStr As String
'    Gp As Gp
'End Type
'Private Type LnxVdt
'    LnxAy() As Lnx
'    Er As TpEr
'End Type
'Private Type FldExpr
'    Fld As String
'    ExprLines As String
'End Type
'Private Enum e_StmtTy
'    e_DrpStmt = 1
'    e_UpdStmt = 2
'    e_SelStmt = 3
'End Enum
'Private Type DrpStmtBrk
'    Tny() As String
'End Type
'Private Type DrpStmtRslt
'    Er As TpEr
'    Drp As DrpStmtBrk
'End Type
'Private Type UpdStmtBrk
'    Tn As String
'End Type
'Private Type SelStmtBrk
'    Tn As String
'End Type
'Private Type StmtBrk_Er
'    Er As TpEr
'    Ty As e_StmtTy
'    Drp As DrpStmtBrk
'    Upd As UpdStmtBrk
'    Sel As SelStmtBrk
'End Type
'Private Type SwBrk
'    Lx As Integer
'    Nm As String
'    OpStr As String
'    TermAy() As String
'End Type
'Private Type SwBrkRslt
'    SwBrk As SwBrk
'    Er As TpEr
'End Type
'Private Type SwBrkAyRslt
'    SwBrkAy() As SwBrk
'    Er As TpEr
'End Type
'Private Type WrkDicRslt
'    WrkDic As Dictionary
'    Er As TpEr
'End Type
'
'Function C_WhyBlkIsEr_MsgAy() As String()
'Dim O$()
'Push O, "The block is error because, it is none of these [RmkBlk SqBlk PrmBlk SwBlk]"
'Push O, "SwBlk is all remark line or SwLin, which is started with ?"
'Push O, "PrmBlk is all remark line or PrmLin, which is started with %"
'Push O, "SqBlk is first non-remark begins with these [sel seldis drp upd] with optionally ?-Pffx"
'Push O, "RmkBlk is all remark lines"
'End Function
'
'Function LinIsTpRmk(A$) As Boolean
'Dim L$: L = LTrim(A)
'If L <> "" Then
'    If HasPfx(L, "--") Then
'        LinIsTpRmk = True
'    End If
'End If
'End Function
'
'Function LnxAy_SwRslt(A() As Lnx, Pm As Dictionary) As SwRslt
''Set A_Pm = Pm
''Dim Brk     As SwBrk:     Brk = LnxAy_SwRslt(A)
''Dim B1      As BrkRslt:    B1 = Z1_TpEr(Brk)
''Dim B2      As BrkRslt:    B2 = Z1_DupNmEr(B1.Rslt)
''Dim W       As WrkDicRslt:  W = Z1_WrkDicRslt(B2.Rslt)
''Dim Er      As TpEr:       Er = W.Er
''                                TpEr_Push Er, B1.Er
''                                TpEr_Push Er, B2.Er
''Dim Sw      As Sw:         Sw = Z1_Sw(W.Rslt)
''With LnxAy_SwRslt
''    .Er = Er
''    .Sw = Sw
''End With
'End Function
'
'Private Function LyIsSw(A$()) As Boolean
'LyIsSw = LyHasMajPfx(A, "?")
'End Function
'
'Function SqFldExprAy_SelPhrase$(A() As FldExpr)
'Dim B() As S1S2
'    Dim B1() As S1S2
'    Dim B2() As S1S2
'    Dim B3() As S1S2
'    Dim B4() As S1S2
'    B2 = SqFldExprAy_SetFldNmPfxIsDot(B1)
'    B3 = SqFldExprAy_RmvEmpExprItm(B2)
'    B4 = SqFldExprAy_RmvFldNmPfxIsQuestionMrk(B3)
'     B = SqFldExprAy_RmvTermWithDot(B4)
'Dim F1$()
'Dim E1$()
'    F1 = S1S2Ay_Sy1(B)
'    E1 = S1S2Ay_Sy2(B)
'Dim F2$()
'Dim E2$()
'    F2 = AyAlignL(F1)
'    E2 = VblAy_AlignAsLy(E1)
'Dim E3$()
''    E3 = Y1__ExprLinesAyTab(E2, 4)
'
'Dim O$()
'    O = S1S2Ay_AddAsLy(SyPair_S1S2Ay(E3, F2), " ")
'SqFldExprAy_SelPhrase = "Select|" & Join(O, ",|")
'End Function
'
'Function SqLin_OptChk$(SqLin)
'If HasPfx(SqLin, "$") Then Exit Function
'Dim T1$: T1 = Lin(SqLin).T1
'Dim O$
'    Select Case RmvPfx(T1, "?")
'    Case "Gp", "Sel", "SelDis", "AndInStrLis", "AndInNbrLis"
'    Case "WhBetStr", "WhBetNbr"
'    Case "Into", "Fm", "Upd"
'    Case Else: O = "Invalid Sql-Phrase-Item"
'    End Select
'If O <> "" Then
'    If FstChr(T1) = "?" Then
'        Dim Ay$(): Ay = SslSy("Sel SelDis Upd AndInStrLis AndInNbrLis WhBetStr WhBetNbr")
'        If Not AyHas(Ay, T1) Then
'            O = Replace("Only these following Sql-Phrase-Item allow [?} as prefix: [*]", "*", JnSpc(Ay))
'        End If
'    End If
'End If
'SqLin_OptChk = O
'End Function
'
'Function SqLnxAy_StmtRslt(A() As Lnx, Pm As Dictionary, Sw As Sw) As StmtRslt
'Dim Ly$(): Ly = LnxAy_Ly(A)
'Dim Ty As e_StmtTy: Ty = SqLy_Ty(Ly)
'Dim IsSkip As Boolean: IsSkip = SqLy_IsSkip(Ly, Ty, Sw.Tbl)
'Dim DrpStmt As StmtRslt: DrpStmt = SqLnxAy_DrpStmtRslt(A, Ty)
'Dim UpdStmt As StmtRslt: UpdStmt = SqLnxAy_UpdStmtRslt(A, Ty, Pm, Sw.Fld)
'Dim SelStmt As StmtRslt: SelStmt = SqLnxAy_SelStmtRslt(A, Ty, Pm, Sw.Fld)
'Dim Stmt$: Stmt = UpdStmt.Stmt + SelStmt.Stmt + DrpStmt.Stmt
'Dim Er As TpEr: Er = TpEr_Add3(UpdStmt.Er, DrpStmt.Er, SelStmt.Er)
'With SqLnxAy_StmtRslt
'    .Er = Er
'    .Stmt = Stmt
'End With
'End Function
'
'Function SqRst_Sel$(A$, EDic As Dictionary)
''Sq:=[S]ql-[T]em[p]late-[C]ontext-For
'Dim Fny$()
'Dim ExprAy$()
'    Fny = SslSy(A)
'Stop
''    ExprAy = DicSelIntoSy(EDic, Fny)
'SqRst_Sel = SqpSel(Fny, ExprAy)
'End Function
'
'Function SqTp_SqTpRslt(SqTp$) As SqTpRslt
'Dim Ly$():                            Ly = SplitCrLf(SqTp)
'Dim G() As Gp:         G = LyGpAy(Ly)
'Dim G1() As Gp:       G1 = GpAy_RmvRmk(G)
'Dim B() As Blk:        B = GpAy_BlkAy(G1)
'
'Dim PmLnxAy() As Lnx:            PmLnxAy = BlkAy_SelLnxAy(B, "PM")
'Dim PmRslt As PmRslt:             PmRslt = PmLnxAy_PmRslt(PmLnxAy)
'Dim Pm As Dictionary:             Set Pm = PmRslt.Pm
'
'Dim SwLnxAy() As Lnx:            SwLnxAy = BlkAy_SelLnxAy(B, "SW")
'Dim SwRslt As SwRslt:             SwRslt = LnxAy_SwRslt(SwLnxAy, Pm)
'Dim Sw As Sw:                         Sw = SwRslt.Sw
'
'Dim SqGpAy() As Gp:               SqGpAy = BlkAy_SqGpAy(B)
'Dim StmtsRslt As StmtsRslt:    StmtsRslt = SqGpAy_StmtsRslt(SqGpAy, Pm, Sw)
'
'Dim ErBlkEr As TpEr:             ErBlkEr = BlkAy_ErBlkEr(B)
'Dim ExcessSwBlkEr As TpEr: ExcessSwBlkEr = BlkAy_ExcessSwBlkEr(B)
'Dim ExcessPmBlkEr As TpEr: ExcessPmBlkEr = BlkAy_ExcessPmBlkEr(B)
'
'Dim Er As TpEr:                       Er = TpErAp_Add6 _
'                                            (PmRslt.Er, _
'                                             SwRslt.Er, _
'                                             StmtsRslt.Er, _
'                                             ExcessPmBlkEr, _
'                                             ExcessSwBlkEr, _
'                                             ErBlkEr)
'
'SqTp_SqTpRslt.MsgLy = TpEr_Ly(Er)
'SqTp_SqTpRslt.Stmts = StmtsRslt.Stmts
'End Function
'
'Private Function BlkAy_ErBlkEr(A() As Blk) As TpEr
''?
'End Function
'
'Private Function BlkAy_ExcessPmBlkEr(A() As Blk) As TpEr
''?
'End Function
'
'Private Function BlkAy_ExcessSwBlkEr(A() As Blk) As TpEr
''?
'End Function
'
'Private Function BlkAy_SelLnxAy(A() As Blk, BlkTyStr$) As Lnx()
'Dim J%
'For J = 0 To BlkUB(A)
'    If A(J).BlkTyStr = BlkTyStr Then BlkAy_SelLnxAy = A(J).Gp.LnxAy: Exit Function
'Next
'End Function
'
'Private Function BlkAy_SqGpAy(A() As Blk) As Gp()
'Dim J%, O() As Gp, M As Gp
'For J = 0 To BlkUB(A)
'    If A(J).BlkTyStr = "SQ" Then
'        M = A(J).Gp
'        GpPush O, M
'    End If
'Next
'BlkAy_SqGpAy = O
'End Function
'
'Private Function BlkSz%(A() As Blk)
'On Error Resume Next
'BlkSz = UBound(A) + 1
'End Function
'
'Private Function BlkUB%(A() As Blk)
'BlkUB = BlkSz(A) - 1
'End Function
'
'Private Function DupNmIxAy_Er(DupNmLx%()) As TpEr
'
'End Function
'
'Private Function GpAy_BlkAy(GpAy() As Gp) As Blk()
'Dim O() As Blk
'Dim U&
'    U = GpUB(GpAy)
'If U = -1 Then Exit Function
'ReDim O(U)
'Dim J%
'For J = 0 To GpUB(GpAy)
'    O(J) = GpBlk(GpAy(J))
'Next
'GpAy_BlkAy = O
'End Function
'
'Private Function GpAy_RmvRmk(A() As Gp) As Gp()
'Dim J%, O() As Gp, M As Gp
'For J = 0 To GpUB(A)
'    M = GpRmvRmk(A(J))
'    If LnxSz(M.LnxAy) > 0 Then
'        GpPush O, M
'    End If
'Next
'GpAy_RmvRmk = O
'End Function
'
'Private Function GpBlkTyStr$(A As Gp)
'Dim Ly$(): Ly = GpLy(A)
'Dim O$
'Select Case True
'Case LyIsPm(Ly): O = "PM"
'Case LyIsSw(Ly): O = "SW"
'Case LyIsRm(Ly): O = "RM"
'Case LyIsSq(Ly): O = "SQ"
'Case Else: O = "ER"
'End Select
'GpBlkTyStr = O
'End Function
'
'Private Function GpBlk(A As Gp) As Blk
'With GpBlk
'    .BlkTyStr = GpBlkTyStr(A)
'    .Gp = A
'End With
'End Function
'
'Private Function GpRmvRmk(A As Gp) As Gp
'Dim B() As Lnx: B = A.LnxAy
'Dim M As Lnx
'Dim J&, O() As Lnx
'For J = 0 To LnxUB(B)
'    M = B(J)
'    If Not LinIsTpRmk(M.Lin) Then
'        LnxPush O, M
'    End If
'Next
'GpRmvRmk.LnxAy = O
'End Function
'
'Private Function LnxAy_ErIx_OfDupKey(A() As Lnx) As Integer()
'
'End Function
'
'Private Function LnxAy_ErIx_OfPfxPercent(A() As Lnx) As Integer()
'
'End Function
'
'Private Function LnxAy_WhByExclErIxAy(A() As Lnx, ErIxAy%()) As String()
'Dim O$(), J%
'For J = 0 To LnxUB(A)
'    If Not AyHas(ErIxAy, J) Then
'        Push O, A(J).Lin
'    End If
'Next
'LnxAy_WhByExclErIxAy = O
'End Function
'
'Private Function LyGpAy(Ly$()) As Gp()
'Dim O() As Gp, J&, LnxAy() As Lnx, M As Lnx
'For J = 0 To UB(Ly)
'    Dim Lin$
'    Lin = Ly(J)
'    If HasPfx(Lin, "==") Then
'        If LnxSz(LnxAy) > 0 Then
'            GpPush O, NewGp(LnxAy)
'        End If
'        Erase LnxAy
'    Else
'        LnxPush LnxAy, NewLnx(J, Lin)
'    End If
'Next
'If LnxSz(LnxAy) > 0 Then
'    GpPush O, NewGp(LnxAy)
'End If
'LyGpAy = O
'End Function
'
'Private Function LyIsPm(A$()) As Boolean
'LyIsPm = LyHasMajPfx(A, "%")
'End Function
'
'Private Function LyIsRm(A$()) As Boolean
'LyIsRm = AyIsEmp(A)
'End Function
'
'Private Function LyIsSq(A$()) As Boolean
'If AyIsEmp(A) Then Exit Function
'Dim L$: L = A(0)
'Dim X$(): X = SslSy("?SEL SEL ?SELDIS SELDIS UPD DRP")
'If HasOneOfPfxIgnCas(L, X) Then LyIsSq = True: Exit Function
'End Function
'
'Private Function PmLnxAy_PmRslt(A() As Lnx) As PmRslt
'Dim ErIx1%()
'Dim ErIx2%()
'    ErIx1 = LnxAy_ErIx_OfDupKey(A)
'    ErIx2 = LnxAy_ErIx_OfPfxPercent(A)
'
'Dim ErIx%()
'    PushAy ErIx, ErIx1
'    PushAy ErIx, ErIx2
'
'Dim Er As TpEr
'Dim ValidatedPmLy$()
'    ValidatedPmLy = LnxAy_WhByExclErIxAy(A, ErIx)
'Dim O As PmRslt
'O.Er = Er
'Stop
''Set O.Pm = LinesDicLy_LinesDic(ValidatedPmLy)
'PmLnxAy_PmRslt = O
'End Function
'
'Private Function SqFldExprAy_RmvEmpExprItm(A() As S1S2) As S1S2()
'Dim O() As S1S2
'Dim F$(), E$()
'Dim J%
'For J = 0 To S1S2_UB(A)
'    If A(J).S2 <> "" Then
'        S1S2_Push O, A(J)
'    End If
'Next
'SqFldExprAy_RmvEmpExprItm = O
'End Function
'
'Private Function SqFldExprAy_RmvFldNmPfxIsQuestionMrk(A() As S1S2) As S1S2()
'Dim O() As S1S2
'O = A
'Dim J%
'For J = 0 To S1S2_UB(O)
'    With O(J)
'        If A(J).S2 <> "" Then
'            If HasPfx(.S1, "?") Then
'                .S1 = RmvFstChr(0.1)
'            End If
'        End If
'    End With
'Next
'SqFldExprAy_RmvFldNmPfxIsQuestionMrk = O
'End Function
'
'Private Function SqFldExprAy_RmvTermWithDot(A() As S1S2) As S1S2()
'Dim J%, O() As S1S2
'O = A
'For J = 0 To S1S2_UB(O)
'    If HasSubStr(O(J).S1, ".") Then
'        O(J).S1 = TakAftRev(O(J).S1, ".")
'        O(J).S1 = RmvPfx(O(J).S1, "?")
'    End If
'Next
'SqFldExprAy_RmvTermWithDot = O
'End Function
'
'Private Function SqFldExprAy_SetFldNmPfxIsDot(A() As S1S2) As S1S2()
'Dim J%, O() As S1S2
'O = A
'For J = 0 To S1S2_UB(O)
'    If FstChr(A(J).S1) = "." Then
'        O(J).S1 = RmvFstChr(O(J).S1)
'        O(J).S2 = O(J).S1
'    End If
'Next
'SqFldExprAy_SetFldNmPfxIsDot = O
'End Function
'
'Private Function SqGpAy_StmtsRslt(SqGpAy() As Gp, Pm As Dictionary, Sw As Sw) As StmtsRslt
'Dim OEr As TpEr
'Dim OStmtAy$()
'    Dim J%
'    For J = 0 To GpUB(SqGpAy)
'        With SqLnxAy_StmtRslt(SqGpAy(J).LnxAy, Pm, Sw)
'            Push OStmtAy, .Stmt
'            TpEr_Push OEr, .Er
'        End With
'    Next
'Dim Z As StmtsRslt
'    Z.Stmts = JnDblCrLf(OStmtAy)
'    Z.Er = OEr
'SqGpAy_StmtsRslt = Z
'End Function
'
'Private Function SqLin_Evl$(SqLin, EDic As Dictionary)
'Dim L$
'    L = SqLin
'Dim QItm$
'Dim IsOpt As Boolean
''    QItm = ParseTerm(L)
'
'Dim Pfx$
'    Select Case QItm
'    Case "Upd":    Pfx = "Update "
'    Case "Set":    Pfx = "|  Set"
'    Case "Sel":    Pfx = "Select"
'    Case "SelDis": Pfx = "Select Distinct"
'    Case "Gp":     Pfx = "|  Group By"
'    Case "Fm":     Pfx = "|  From "
'    Case "Left":   Pfx = "|  Left Join "
'    Case "Jn":     Pfx = "|  Join "
'    Case _
'        "WhEDic", _
'        "WhInStrLis", _
'        "WhInNbrLis", _
'        "WhBetStr", _
'        "WhBetNbr":
'                   Pfx = "|  Where "
'    Case _
'        "AndEDic", _
'        "AndInStrLis", _
'        "AndInNbrLis", _
'        "AndBetStr", _
'        "AndBetNbr":
'                   Pfx = "|  And "
'    Case Else
'        Stop
'    End Select
'Dim Rst$
'    Select Case QItm
'    Case _
'        "Upd", _
'        "Fm", _
'        "Left", _
'        "Jn", _
'        "WhEDic", _
'        "AndEDic"
'                        Rst = L
'    Case _
'        "Sel", _
'        "SelDis"
'                        Rst = SqRst_Sel(Rst, EDic)    ' Sq = Sql-Phrase-Itm-Context
'    Case "WhInStrLis":  Rst = SqRst_InLis(Rst, EDic, IsStr:=True)
'    Case "WhInNbrLis":  Rst = SqRst_InLis(Rst, EDic, IsStr:=False)
'    Case "WhBetStr":    Rst = SqRst_Bet(Rst, EDic, IsStr:=True)
'    Case "WhBetNbr":    Rst = SqRst_Bet(Rst, EDic, IsStr:=False)
'    Case "Gp":          Rst = SqRst_Gp(Rst, EDic)
'    Case Else
'        Stop
'    End Select
'SqLin = Pfx & Rst
'
'End Function
'
'Private Function SqLnxAy_DrpStmtRslt(SqLnxAy() As Lnx, Ty As e_StmtTy) As StmtRslt
'If Ty = e_DrpStmt Then Exit Function
'Dim TnLvs$
'With SqLnxAy_DrpStmtRslt
'    .Stmt = TblNms(TnLvs).DrpStmts
'End With
'End Function
'
'Private Function SqLnxAy_SelStmtRslt(SqLnxAy() As Lnx, Ty As e_StmtTy, Pm As Dictionary, FldSw As Dictionary) As StmtRslt
'If Ty = e_SelStmt Then Exit Function
'End Function
'
'Private Function SqLnxAy_UpdStmtRslt(SqLnxAy() As Lnx, Ty As e_StmtTy, Pm As Dictionary, FldSw As Dictionary) As StmtRslt
'If Ty <> e_UpdStmt Then Exit Function
'End Function
'
'Private Function SqLy_IsSkip(A$(), Ty As e_StmtTy, TblSw As Dictionary) As Boolean
'Dim T$: T = SqLy_TarTn(A, Ty)
'SqLy_IsSkip = TblSw.Exists(T)
'End Function
'
'Private Function SqLy_TarTn$(A$(), Ty As e_StmtTy)
'
'End Function
'
'Private Function SqLy_Ty(Ly$()) As e_StmtTy
'
'End Function
'
'Private Function SqRst_Bet$(A$, EDic As Dictionary, IsStr As Boolean)
'Dim ETerm$, T1$, T2$
'Dim C$, E$
'    Const C1$ = "? Between '?' and '?'"
'    Const C2$ = "? Between ? and ?"
'    C = IIf(IsStr, C1, C2)
'SqRst_Bet = FmtQQ(C, E, T1, T2)
'End Function
'
'Private Function SqRst_Gp$(A$, EDic As Dictionary)
'Dim ExprAy$(), Ay$()
'Stop
''    ExprAy = DicSelIntoSy(EDic, Ay)
'SqRst_Gp = SqpGp(ExprAy)
'End Function
'
'Private Function SqRst_InLis$(A$, EDic As Dictionary, IsStr As Boolean)
'Dim ETerm$, TermAy$()
'Ass EDic.Exists(ETerm)
'Dim L$, E$
'    If IsStr Then L = JnQSngComma(TermAy) Else L = JnComma(TermAy)
'SqRst_InLis = FmtQQ("? in ( ? )", E, L)
'End Function
'
'Private Function SwBrkAy_DupNmEr(A() As SwBrk) As TpEr
'Dim Ny$(): Ny = SwBrkAy_Ny(A)
'Dim DupNy$(): DupNy = AyDupAy(Ny)
'Dim DupNmLx%(): DupNmLx = SwBrkAy_DupNmIx(A, DupNy)
'Dim OEr As TpEr: OEr = DupNmIxAy_Er(DupNmLx)
'Dim ORslt As SwBrk: ORslt = SwBrkAy_WhExclByIxAy(A, DupNmLx)
'With SwBrkAy_DupNmEr
'    '.OEr
'    '.Rslt = ORslt
'End With
'End Function
'
'Private Function SwBrkAy_DupNmIx(A() As SwBrk, DupNy$()) As Integer()
'
'End Function
'
'Private Function SwBrkAy_Ny(A() As SwBrk) As String()
'Dim O$(), J%
'For J = 0 To SwBrk_UB(A)
'    Push O, A(J).Nm
'Next
'SwBrkAy_Ny = O
'End Function
'
'Private Function SwBrkAy_Rslt(A() As SwBrk, Pm As Dictionary) As SwBrkAyRslt
'Dim O() As SwBrk
'Dim J%, OEr As TpEr, M As TpEr
'For J = 0 To SwBrk_Sz(A)
'    M = SwBrk_Er(A(J), Pm)
'    With M
'        If .N = 0 Then
'            SwBrk_Push O, A(J)
'        Else
'            TpEr_Push OEr, M
'        End If
'    End With
'Next
'With SwBrkAy_Rslt
'    .Er = OEr
'    .SwBrkAy = O
'End With
'End Function
'
'Private Function SwBrkAy_WhExclByIxAy(A() As SwBrk, Lx%()) As SwBrk
'Dim O As SwBrk
'Dim J%
'For J = 0 To SwBrk_UB(A)
'    If Not AyHas(Lx, A(J).Lx) Then
'        SwBrk_Push A, A(J)
'    End If
'Next
'SwBrkAy_WhExclByIxAy = O
'End Function
'
'Private Function SwBrkAy_WrkDicRslt(A() As SwBrk, Pm As Dictionary) As WrkDicRslt
'Dim Sy$()
'Dim SomLinEvaluated As Boolean
'Dim I%, J%
'Dim OSw As New Dictionary
'SomLinEvaluated = True
'While SomLinEvaluated
'    I = I + 1
'    If I > 1000 Then Stop
'    SomLinEvaluated = False
'    For J = 0 To SwBrk_Sz(A)
'        With SwBrk_SomBool(A(J), Pm, OSw)
'            If .Som Then
'                SomLinEvaluated = True
'                OSw.Add A(J).Nm, .Bool         '<==
'            End If
'        End With
'    Next
'Wend
'Dim Er As TpEr
''If Z.Count <> Sz(AyRmvEmp(A)) Then Stop
''With SwRslt
''    .Er = Er
''    .Sw = Sw
''End With
'End Function
'
'Private Function SwBrk_AndOrLinEr(A As SwBrk, Pm As Dictionary) As TpEr
'If Not BoolOpStr_IsAndOr(A.OpStr) Then Exit Function
'Dim O As TpEr
'Dim NTerm%
'Dim Lx%
'    NTerm = Sz(A.TermAy)
'    Lx = A.Lx
'
'If NTerm < 2 Then O = NewTpEr(Lx, "For OR|AND, must have 2 or more operands"): GoTo Ext
'
'Dim T1$: T1 = A.TermAy(0)
'Dim T2$: T2 = A.TermAy(1)
'
'Select Case FstChr(T1)
'    Case "%"
'        If Not Pm.Exists(T1) Then
'            O = NewTpEr(Lx, "For OR|AND, first term must be found in Pm")
'            GoTo Ext
'        End If
'    Case Else
'        O = NewTpEr(Lx, "For OR|AND, first operand must begin with %")
'        GoTo Ext
'    End Select
'
'Select Case FstChr(T2)
'    Case "%"
'        If Not Pm.Exists(T2) Then
'            O = NewTpEr(Lx, "For EQ|NE, second operand not found in Pm")
'            GoTo Ext
'        End If
'    Case "?"
'        O = NewTpEr(Lx, "For EQ|NE, second operand cannot begin with ?")
'        GoTo Ext
'    Case "*"
'        If UCase(T1) <> "*BLANK" Then
'            O = NewTpEr(Lx, "For AND|OR, second operand can be *BLANK, but nothing else begin with *")
'            GoTo Ext
'        End If
'    End Select
'Ext: SwBrk_AndOrLinEr = O
'End Function
'
'Private Function SwBrk_SomBool(A As SwBrk, Pm As Dictionary, Sw As Dictionary) As SomBool
'If Sw.Exists(A.Nm) Then Exit Function
'Dim Ay$(): Ay = A.TermAy
'Dim Z As SomBool
'Select Case A.OpStr
'Case "OR":  Z = SwTermAy_SomBool(Ay, "OR", Pm, Sw)
'Case "AND": Z = SwTermAy_SomBool(Ay, "AND", Pm, Sw)
'Case "NE":  Z = SwT1T2_SomBool(Ay(0), Ay(1), "NE", Pm, Sw)
'Case "EQ":  Z = SwT1T2_SomBool(Ay(0), Ay(1), "EQ", Pm, Sw)
'Case Else: Stop
'End Select
'If Not Z.Som Then Exit Function
'SwBrk_SomBool = SomBool(Z.Bool)
'End Function
'
'Private Function SwBrk_EqNeLinEr(A As SwBrk) As TpEr
'If Not BoolOpStr_IsEqNe(A.OpStr) Then Exit Function
'Dim NTerm%, TermAy$()
'Dim Lx%
'With A
'    Lx = .Lx
'    TermAy = .TermAy
'    NTerm = Sz(TermAy)
'End With
'Dim O As TpEr
'If NTerm <> 2 Then
'    O = NewTpEr(Lx, "When 2nd-Term (Operator) is [AND OR], only 2 terms are allowed")
'    GoTo Ext
'End If
'Ext:
'    SwBrk_EqNeLinEr = O
'End Function
'
'Private Function SwBrk_Er(A As SwBrk, Pm As Dictionary) As TpEr
'Dim O As TpEr
'TpEr_Push O, SwBrk_NmEr(A)
'TpEr_Push O, SwBrk_OpEr(A)
'TpEr_Push O, SwBrk_PfxEr(A)
'TpEr_Push O, SwBrk_EqNeLinEr(A)
'TpEr_Push O, SwBrk_AndOrLinEr(A, Pm)
'SwBrk_Er = O
'End Function
'
'Private Function SwBrk_NmEr(A As SwBrk) As TpEr
'If A.Nm = "" Then
'    SwBrk_NmEr = NewTpEr(A.Lx, "The line has no name")
'End If
'End Function
'
'Private Function SwBrk_OpEr(A As SwBrk) As TpEr
'If BoolOpStr_IsVdt(A.OpStr) Then Exit Function
'SwBrk_OpEr = NewTpEr(A.Lx, "Invalid operator.  Valid operation [NE EQ AND OR]")
'End Function
'
'Private Function SwBrk_PfxEr(A As SwBrk) As TpEr
'If FstChr(A.Nm) <> "?" Then
'    SwBrk_PfxEr = NewTpEr(A.Lx, "First char must be [?]")
'End If
'End Function
'
'Private Sub SwBrk_Push(O() As SwBrk, A As SwBrk)
'Dim N%: N = SwBrk_Sz(O)
'ReDim Preserve O(N)
'O(N) = A
'End Sub
'
'Private Function SwBrk_Sz%(A() As SwBrk)
'On Error Resume Next
'SwBrk_Sz = UBound(A) + 1
'End Function
'
'Private Function SwBrk_TermEr(A As SwBrk, PmNmSet As Dictionary) As TpEr
'Dim O0$(), O1$(), O2$(), I
'Dim FldNmSet As Dictionary
'Dim TermAy$()
'Dim O As TpEr
'Dim Lx%
'For Each I In TermAy
'    If HasPfx(CStr(I), "?") Then
'        If Not FldNmSet.Exists(I) Then Push O0, I
'    ElseIf HasPfx(CStr(I), "%?") Then
'        If Not PmNmSet.Exists(I) Then Push O1, I
'    Else
'        Push O2, I
'    End If
'Next
'Dim B$, C$
'If Not AyIsEmp(O0) Then C = FmtQQ("[?] must be found in Switch", JnSpc(O0))
'If Not AyIsEmp(O1) Then C = FmtQQ("[?] must be found in Pm", JnSpc(O1))
'If Not AyIsEmp(O2) Then B = FmtQQ("[?] must begin with [ ? | %? ]", JnSpc(O1))
'Dim Sy$()
'    PushNonEmp Sy, C
'    PushNonEmp Sy, B
'If Sz(Sy) > 0 Then
'    O = NewTpEr(Lx, JnCrLf(Sy))
'End If
'
'End Function
'
'Private Function SwBrk_UB%(A() As SwBrk)
'SwBrk_UB = SwBrk_Sz(A) - 1
'End Function
'
'Private Function SwLnxAy_SwBrkAy(A() As Lnx) As SwBrk()
'Dim J%, O() As SwBrk
'Dim U%: U = LnxUB(A)
'If U = -1 Then Exit Function
'ReDim O(U)
'For J = 0 To U
'    O(J) = SwLnx_SwBrk(A(J))
'Next
'SwLnxAy_SwBrkAy = O
'End Function
'
'Private Function SwLnx_SwBrk(A As Lnx) As SwBrk
'Dim Z As SwBrk
''If SrcLin_IsRmk(SwLin) Then Exit Function 'assume SwLnx has remark removed
'    Dim L$, NTerm%, TermAy$()
'    With Lin(A.Lin)
'        Z.Nm = .ShiftTerm
'        Z.OpStr = UCase(.ShiftTerm)
'        Z.TermAy = SslSy(.Lin)
'        Z.Lx = A.Lx
'    End With
'SwLnx_SwBrk = Z
'End Function
'
'Private Function SwT1T2_SomBool(T1$, T2$, EqNeOpStr$, Pm As Dictionary, Sw As Dictionary) As SomBool
'Dim S1$, S2$
'    With SwTerm_VarOpt(T1, Pm, Sw)
'        If Not .Som Then Exit Function
'        S1 = .V
'    End With
'    With SwTerm_VarOpt(T2, Pm, Sw)
'        If Not .Som Then Exit Function
'        S2 = .V
'    End With
'Dim O As SomBool
'Select Case EqNeOpStr
'Case "EQ": O = SomBool(S1 = S2)
'Case "NE": O = SomBool(S1 <> S2)
'Case Else: Stop
'End Select
'SwT1T2_SomBool = O
'End Function
'
'Private Function SwTerm_VarOpt(A, Pm As Dictionary, Sw As Dictionary) As VarOpt
''switch-term begins with % or ? or it is *Blank.  % is for parameter & ? is for switch
''  If %, it will evaluated to str by lookup from Pm
''        if not exist in {Pm}, stop, it means the validation fail to remove this term
''  If ?, it will evaluated to bool by lookup from Sw
''        if not exist in {Sw}, return None
''  Otherwise, just return SomVar(A)
'Dim O As VarOpt
'    Select Case FstChr(A)
'    Case "?"
'        If Not Sw.Exists(A) Then Exit Function
'        O = SomVar(Sw(A))
'    Case "%"
'        If Not Pm.Exists(A) Then
'            Stop ' it means the validation fail to remove this term
'        End If
'        O = SomVar(Pm(A))
'    Case "*"
'        If A <> "*Blank" Then Stop ' it means the validation fail to remove this term
'        O = SomVar("")
'    Case Else
'        O = SomVar(A)
'    End Select
'SwTerm_VarOpt = O
'End Function
'
'Private Function SwTermAy_SomBool(A$(), Op As e_BoolAyOp, Pm As Dictionary, Sw As Dictionary) As SomBool
'Dim B As New Bools
'    Dim I
'    For Each I In A
'        With SwTerm_VarOpt(I, Pm, Sw)
'            If Not .Som Then Exit Function
'            B.Push CBool(.V)
'        End With
'    Next
'SwTermAy_SomBool = SomBool(B.Val(Op))
'End Function
'
'Private Function SwWrkDic_Sw(A As Dictionary) As Sw
'
'End Function
'
'
'Private Function TpEr_Add3(A1 As TpEr, A2 As TpEr, A3 As TpEr) As TpEr
'Dim O As TpEr
'TpEr_Add3 = O
'End Function
'
'Private Function ZZGpAy() As Gp()
''ZZGpAy = LyGpAy(ZZSqTpLy)
'End Function
'
'Private Function ZZMd() As CodeModule
''Set ZZMd = Md("SqTpSw")
'End Function
'
'Private Function ZZPm() As Dictionary
'Stop
''Set ZZPm = NewLyDic(ZZPmLy)
'End Function
'
'Private Function ZZPmLy() As String()
'ZZPmLy = MdResLy(ZZMd, "PmLy")
'End Function
'
'Private Sub ZZResPmLy()
''sldkfj skldjf '
'' skdfjl
''sdfl sdkfl
'End Sub
'
'Private Sub ZZResSqTp()
''-- Rmk: -- is remark
''-- %XX: is prmDicLin
''-- %?XX: is switchPrm, it value must be 0 or 1
''-- ?XX: is switch line
''-- SwitchLin: is ?XXX [OR|AND|EQ|NE] [SwPrm_OR_AND|SwPrm_EQ_NE]
''-- SwPrm_OR_AND: SwTerm ..
''-- SwPrm_EQ_NE:  SwEQ_NE_T1 SwEQ_NE_T2
''-- SwEQ_NE_T1:
''-- SwEQ_NE_T2:
''-- SwTerm:     ?XX|%?XX     -- if %?XX, its value only 1 or 0 is allowed
''-- Only one gp of %XX:
''-- Only one gp of ?XX:
''-- All other gp is sql-statement or sql-statements
''-- sql-statments: Drp xxx xxx
''-- sql-statment: [sel|selDis|upd|into|fm|whBetStr|whBetNbr|whInStrLis|whInNbrLis|andInNbrLis|andInStrLis|gp|jn|left|expr]
''-- optional: Whxxx and Andxxx can have ?-pfx becomes: ?Whxxx and ?Andxxx.  The line will become empty
''==============================================
''Drp Tx TxMbr MbrDta Div Sto Crd Cnt Oup MbrWs
''=============================================
''-- %? means switch, value must be 0 or 1
''%?BrkMbr 0
''%?BrkMbr 0
''%?BrkMbr 0
''%?BrkSto 0
''%?BrkCrd 0
''%?BrkDiv 0
''-- %XXX means txt and optional, allow, blank
''%SumLvl  Y
''%?MbrEmail 1
''%?MbrNm    1
''%?MbrPhone 1
''%?MbrAdr   1
''-- %% mean compulasary
''%%DteFm 20170101
''%%DteTo 20170131
''%LisDiv 1 2
''%LisSto
''%LisCrd
''%CrdExpr ...
''%CrdExpr ...
''%CrdExpr ...
''============================================
''-- EQ & NE t1 only TxtPm is allowed
''--         t2 allow TxtPm, *BLANK, and other text
''?LvlY    EQ %SumLvl Y
''?LvlM    EQ %SumLvl M
''?LvlW    EQ %SumLvl W
''?LvlD    EQ %SumLvl D
''?Y       OR ?LvlD ?LvlW ?LvlM ?LvlY
''?M       OR ?LvlD ?LvlW ?LvlM
''?W       OR ?LvlD ?LvlW
''?D       OR ?LvlD
''?Dte     OR ?LvlD
''?Mbr     OR %?BrkMbr
''?MbrCnt  OR %?BrkMbr
''?Div     OR %?BrkDiv
''?Sto     OR %?BrkSto
''?Crd     OR %?BrkCrd
''?#SEL#Div NE %LisDiv *blank
''?#SEL#Sto NE %LisSto *blank
''?#SEL#Crd NE %LisCrd *blank
''============================================= #Tx
''sel  ?Crd ?Mbr ?Div ?Sto ?Y ?M ?W ?WD ?D ?Dte Amt Qty Cnt
''into #Tx
''fm   SalesHistory
''wh   bet str    %%DteFm %%DteTo
''?and in  strlis Div %LisDiv
''?and in  strlis Sto %LisSto
''?and in  nbrlis Crd %LisCrd
''?gp  ?Crd ?Mbr ?Div ?Sto ?Crd ?Y ?M ?W ?WD ?D ?Dte
''$Crd %CrdExpr
''$Mbr JCMCode
''$Sto
''$Y
''$M
''$W
''$WD
''$D
''$Dte
''$Amt Sum(SHAmount)
''$Qty Sum(SHQty)
''$Cnt Count(SHInvoice+SHSDate+SHRef)
''============================================= #TxMbr
''selDis  Mbr
''fm      #Tx
''into    #TxMbr
''============================================= #MbrDta
''sel   Mbr Age Sex Sts Dist Area
''fm    #TxMbr x
''jn    JCMMember a on x.Mbr = a.JCMMCode
''into  #MbrDta
''$Mbr  x.Mbr
''$Age  DATEDIFF(YEAR,CONVERT(DATETIME ,x.JCMDOB,112),GETDATE())
''$Sex  a.JCMSex
''$Sts  a.JCMStatus
''$Dist a.JCMDist
''$Area a.JCMArea
''==-=========================================== #Div
''?sel Div DivNm DivSeq DivSts
''fm   Division
''into #Div
''?wh in strLis Div %LisDiv
''$Div Dept + Division
''$DivNm LongDies
''$DivSeq Seq
''$DivSts Status
''============================================ #Sto
''?sel Sto StoNm StoCNm
''fm   Location
''into #Sto
''?wh in strLis Loc %LisLoc
''$Sto
''$StoNm
''$StoCNm
''============================================= #Crd
''?sel        Crd CrdNm
''fm          Location
''into        #Crd
''?wh in nbrLis Crd %LisCrd
''$Crd
''$CrdNm
''============================================= #Oup
''sel  ?Crd ?CrdNm ?Mbr ?Age ?Sex ?Sts ?Dist ?Area ?Div ?DivNm ?Sto ?StoNm ?StoCNm ?Y ?M ?W ?WD ?D ?Dte Amt Qty Cnt
''into #Oup
''fm   #Tx x
''left #Crd a on x.Crd = a.Crd
''left #Div b on x.Div = b.Div
''left #Sto c on x.Sto = c.Sto
''left #MbrDta d on x.Mbr = d.Mbr
''wh   JCMCode in (Select Mbr From #TxMbr)
''============================================ #Cnt
''sel ?MbrCnt RecCnt TxCnt Qty Amt
''into #Cnt
''fm  #Tx
''$MbrCnt Count(Distinct Mbr)
''$RecCnt Count(*)
''$TxCnt  Sum(TxCnt)
''$Qty    Sum(Qty)
''$Amt    Sum(Amt)
''============================================
''--
''============================================
''df eror fs--
''============================================
''-- EQ & NE t1 only TxtPm is allowed
''--         t2 allow TxtPm, *BLANK, and other text
''?LvlY    EQ %SumLvl Y
''?LvlM    EQ %SumLvl M
''?LvlW    EQ %SumLvl W
''?LvlD    EQ %SumLvl D
''?Y       OR ?LvlD ?LvlW ?LvlM ?LvlY`
'End Sub
'
'Private Sub ZZResSwLy()
''sdfsdf
''sdfsdfa
'End Sub
'
'Private Sub ZZResWhereTp()
'
'' wh * Bet * Str|Nbr
'' wh * In  * Str|Nbr * Lis
'' and * Bet * Str|Nbr
'' and * In  * Str|Nbr * Lis
'' ==> wh|and * ( ( Bet * Str|Nbr) | (In * Str|Nbr * Lis) )
'' ==> wh|and * Bet|In * Str|Nbr * Lis|_
'' ==> wh|and * Bet|In * Str|Nbr
'End Sub
'
'Private Function ZZSqTp$()
'Static X$
''If X = "" Then X = MdResStr(Md("W01SqTp"), "SqTp")
'ZZSqTp = X
'End Function
'
'Private Function ZZSqTpLy() As String()
'ZZSqTpLy = SplitCrLf(ZZSqTp)
'End Function
'
'Private Function ZZSwBrkAy() As SwBrk()
'ZZSwBrkAy = SwLnxAy_SwBrkAy(ZZSwLnxAy)
'End Function
'
'Private Function ZZSwBrkAyNoEr() As SwBrk()
''Dim B1 As BrkRslt: B1 = Z1_TpEr(ZZSwBrk)
''Dim B2 As BrkRslt: B2 = Z1_DupNmEr(B1.Rslt)
''ZZSwBrkAyNoEr = B2.Rslt
'End Function
'
'Private Function ZZSwLnxAy() As Lnx()
''Dim Ly$(): Ly = MdResLy(Md("SqTpSw"), "SwLy")
''ZZSwLnxAy = LyLnxAy(Ly)
'End Function
'
'Private Function ZZSwLy() As String()
'ZZSwLy = MdResLy(ZZMd, "SwLy")
'End Function
'
'Private Sub FmtSql__Tst()
'Dim Tp$: Tp = "Select" & _
'"|{?Sel}" & _
'"|    {ECrd} Crd," & _
'"|    {EAmt} Amt," & _
'"|    {EQty} Qty," & _
'"|    {ECnt} Cnt," & _
'"|  Into #Tx" & _
'"|  From SaleHistory" & _
'"|  Where SHDate Between '{PFm}' and '{PTo}'" & _
'"|    And {EDiv} in ({InDiv})" & _
'"|  Group By" & _
'"|{?Gp}" & _
'"|?M   {ETxM}," & _
'"|?W   {ETxW}," & _
'"|?D   {ETxD}"
''SR_ = Sales Report
'Const ETxWD$ = _
'"CASE WHEN TxWD1 = 1 then 'Sun'" & _
'"|ELSE WHEN TxWD1 = 2 THEN 'Mon'" & _
'"|ELSE WHEN TxWD1 = 3 THEN 'Tue'" & _
'"|ELSE WHEN TxWD1 = 4 THEN 'Mon'" & _
'"|ELSE WHEN TxWD1 = 5 THEN 'Thu'" & _
'"|ELSE WHEN TxWD1 = 6 THEN 'Fri'" & _
'"|ELSE WHEN TxWD1 = 7 THEN 'Sat'" & _
'"|ELSE Null" & _
'"|END END END END END END END"
'Dim D As New Dictionary
'With D
'    .Add "ECrd", "Line-1|Line-2"
'    .Add "EAmt", "Sum(SHTxDate)"
'
'End With
'Dim Act$: 'Act = FmtSql(Tp, D)
'Dim Exp$: Exp = ""
'Ass Act = Exp
'End Sub
'
'Private Sub LnxAy_SwRslt__Tst()
'Dim Act As SwRslt
'Act = LnxAy_SwRslt(ZZSwLnxAy, ZZPm)
'Stop
'End Sub
'
'Private Sub SqTp_SqTpRslt__Tst()
'Dim A As SqTpRslt: A = SqTp_SqTpRslt(ZZSqTp)
'Stop
'End Sub
'
'Private Sub ZZSqTp__Tst()
'Debug.Print ZZSqTp
'End Sub
'

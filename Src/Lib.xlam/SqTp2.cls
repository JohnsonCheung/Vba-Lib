VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqTp2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const C_BlkTyLvs$ = "ER PM SW SQ RM"
Const Msg_Sq_1_NotInEDic = "These items not found in ExprDic [?]"
Const Msg_Sq_1_MustBe1or0 = "For %?xxx, 2nd term must be 1 or 0"
Const QItmLvs0$ = "sel selDis gp"
Const C_LvsOf_1$ = "into fm left jn" ' QItm = Sql-Phrase-Itm
Const C_LvsOf_2$ = "upd set"
Const C_LvsOf_3$ = "whExpr whBetStr whBetNbr whInStrLis whInNbrLis"
Const QItmLvs4$ = "andExpr andBetStr andBetNbr andInStrLis andInNbrLis"
'Const QItmLvs$ = QItmLvs0 & QItmLvs1 & QItmLvs2 & QItmLvs3 & QItmLvs4
Const OptionalQItm$ = "sel selDis wh* and*"
'============
Private Type Blk
    BlkTyStr As String
    Lnxs As Lnx1s
End Type
Private Type LnxVdt
    LnxAy() As Lnx
    Er As TpEr
End Type
Private Type FldExpr
    Fld As String
    ExprLines As String
End Type
Private Enum e_StmtTy
    e_DrpStmt = 1
    e_UpdStmt = 2
    e_SelStmt = 3
End Enum
Private Type DrpStmtBrk
    Tny() As String
End Type
Private Type DrpStmtRslt
    Er As TpEr
    Drp As DrpStmtBrk
End Type
Private Type UpdStmtBrk
    Tn As String
End Type
Private Type SelStmtBrk
    Tn As String
End Type
Private Type StmtBrk_Er
    Er As TpEr
    Ty As e_StmtTy
    Drp As DrpStmtBrk
    Upd As UpdStmtBrk
    Sel As SelStmtBrk
End Type
Private Type SwBrk
    Lx As Integer
    Nm As String
    OpStr As String
    TermAy() As String
End Type
Private Type SwBrkRslt
    SwBrk As SwBrk
    Er As TpEr
End Type
Private Type SwBrkAyRslt
    SwBrkAy() As SwBrk
    Er As TpEr
End Type
Private Type WrkDicRslt
    WrkDic As Dictionary
    Er As TpEr
End Type
Private B_SqTp$
Private Function C_WhyBlkIsEr_MsgAy() As String()
Dim O$()
Push O, "The block is error because, it is none of these [RmkBlk SqBlk PrmBlk SwBlk]"
Push O, "SwBlk is all remark line or SwLin, which is started with ?"
Push O, "PrmBlk is all remark line or PrmLin, which is started with %"
Push O, "SqBlk is first non-remark begins with these [sel seldis drp upd] with optionally ?-Pffx"
Push O, "RmkBlk is all remark lines"
End Function

Private Function LinIsTpRmk(A$) As Boolean
Dim L$: L = LTrim(A)
If L <> "" Then
    If HasPfx(L, "--") Then
        LinIsTpRmk = True
    End If
End If
End Function

Private Function SqFldExprAy_SelPhrase$(A() As FldExpr)
Dim B() As S1S2
    Dim B1() As S1S2
    Dim B2() As S1S2
    Dim B3() As S1S2
    Dim B4() As S1S2
    B2 = SqFldExprAy_SetFldNmPfxIsDot(B1)
    B3 = SqFldExprAy_RmvEmpExprItm(B2)
    B4 = SqFldExprAy_RmvFldNmPfxIsQuestionMrk(B3)
     B = SqFldExprAy_RmvTermWithDot(B4)
Dim F1$()
Dim E1$()
    F1 = S1S2Ay_Sy1(B)
    E1 = S1S2Ay_Sy2(B)
Dim F2$()
Dim E2$()
    F2 = AyAlignL(F1)
    E2 = VblAy_AlignAsLy(E1)
Dim E3$()
'    E3 = Y1__ExprLinesAyTab(E2, 4)

Dim O$()
    O = S1S2Ay_AddAsLy(SyPair_S1S2Ay(E3, F2), " ")
SqFldExprAy_SelPhrase = "Select|" & Join(O, ",|")
End Function

Private Function SqLin_OptChk$(SqLin)
If HasPfx(SqLin, "$") Then Exit Function
Dim T1$: T1 = Lin(SqLin).T1
Dim O$
    Select Case RmvPfx(T1, "?")
    Case "Gp", "Sel", "SelDis", "AndInStrLis", "AndInNbrLis"
    Case "WhBetStr", "WhBetNbr"
    Case "Into", "Fm", "Upd"
    Case Else: O = "Invalid Sql-Phrase-Item"
    End Select
If O <> "" Then
    If FstChr(T1) = "?" Then
        Dim Ay$(): Ay = LvsSy("Sel SelDis Upd AndInStrLis AndInNbrLis WhBetStr WhBetNbr")
        If Not AyHas(Ay, T1) Then
            O = Replace("Only these following Sql-Phrase-Item allow [?} as prefix: [*]", "*", JnSpc(Ay))
        End If
    End If
End If
SqLin_OptChk = O
End Function

Private Function SqLnxs_StmtRslt(A As Lnx1s, Pm As Dictionary, Sw As Sw) As StmtRslt1
Dim Ly$(): Ly = A.Ly
Dim Ty As e_StmtTy: Ty = SqLy_Ty(Ly)
Dim IsSkip As Boolean: IsSkip = SqLy_IsSkip(Ly, Ty, Sw.Tbl)
Dim DrpStmt As StmtRslt1: DrpStmt = SqLnxs_DrpStmtRslt(A, Ty)
Dim UpdStmt As StmtRslt1: UpdStmt = SqLnxs_UpdStmtRslt(A, Ty, Pm, Sw.Fld)
Dim SelStmt As StmtRslt1: SelStmt = SqLnxs_SelStmtRslt(A, Ty, Pm, Sw.Fld)
Dim Stmt$: Stmt = UpdStmt.Stmt + SelStmt.Stmt + DrpStmt.Stmt
Dim E$(): E = Ly0Ap_Ly(UpdStmt.ErLy, DrpStmt.ErLy, SelStmt.ErLy)
With SqLnxs_StmtRslt
    .ErLy = E
    .Stmt = Stmt
End With
End Function

Private Function SqRst_Sel$(A$, EDic As Dictionary)
'Sq:=[S]ql-[T]em[p]late-[C]ontext-For
Dim Fny$()
Dim ExprAy$()
    Fny = LvsSy(A)
Stop
'    ExprAy = DicSelIntoSy(EDic, Fny)
SqRst_Sel = SqpSel(Fny, ExprAy)
End Function

Friend Property Get Init(Optional SqTpStr) As SqTp2
If IsNonBlankStr(SqTpStr) Then
    B_SqTp = SqTpStr
Else
    B_SqTp = ZZSqTp
End If
Set Init = Me
End Property

Property Get Validate() As SqTpRslt
Dim L$():                              L = SplitCrLf(B_SqTp)
Dim L1 As Lnx1s:                  Set L1 = Ly(L).Lnxs.RmvDDRmk
Dim L2() As Lnx1s:                    L2 = L1.LnxsAy("==")
Dim B As Blk1s:                    Set B = LnxsAy_BlkAy(L2)

Dim PmRslt As PmRslt:      Set PmRslt = B.PmRslt
Dim Pm As Dictionary:             Set Pm = PmRslt.Pm

Dim SwRslt As SwRslt:         Set SwRslt = B.SwRslt
Dim Sw As Sw:                     Set Sw = SwRslt.Sw

Dim StmtsRslt As StmtsRslt:    StmtsRslt = B.StmtsRslt
Dim Stmts$:                        Stmts = StmtsRslt.Stmts

Dim Tp$: '                            Tp = S
Dim Er1$():                            Er1 = Ly0Ap_Ly( _
                                             B.ErLy, _
                                             PmRslt.ErLy, _
                                             SwRslt.ErLy, _
                                             StmtsRslt.ErLy)
                                             
Dim O As New SqTpRslt
O.Er = Er1
O.Stmts = Stmts
O.Tp = Tp
Set Validate = O
End Property


Private Function LnxsAy_BlkAy(A() As Lnx1s) As Blk1s
Dim Ay() As Blk1
Dim U&
    U = UB(A)
If U = -1 Then Exit Function
ReDim Ay(U)
Dim J%
For J = 0 To U
    Dim M As New Blk1
    Set Ay(J) = M.Init(A(J))
Next
Dim O As New Blk1s
Set LnxsAy_BlkAy = O.Init(Ay)
End Function



Private Function LnxAy_ErIx_OfDupKey1(A() As Lnx1) As Integer()

End Function

Private Function LnxAy_ErIx_OfPfxPercent(A() As Lnx1) As Integer()

End Function

Private Function LnxAy_WhByExclErIxAy(A() As Lnx1, ErIxAy%()) As String()
Dim O$(), J%
For J = 0 To UB(A)
    If Not AyHas(ErIxAy, J) Then
        Push O, A(J).Lin
    End If
Next
LnxAy_WhByExclErIxAy = O
End Function

Private Function PmLnxAy_PmRslt(A() As Lnx1) As PmRslt
End Function

Private Function SqFldExprAy_RmvEmpExprItm(A() As S1S2) As S1S2()
Dim O() As S1S2
Dim F$(), E$()
Dim J%
For J = 0 To S1S2_UB(A)
    If A(J).S2 <> "" Then
        S1S2_Push O, A(J)
    End If
Next
SqFldExprAy_RmvEmpExprItm = O
End Function

Private Function SqFldExprAy_RmvFldNmPfxIsQuestionMrk(A() As S1S2) As S1S2()
Dim O() As S1S2
O = A
Dim J%
For J = 0 To S1S2_UB(O)
    With O(J)
        If A(J).S2 <> "" Then
            If HasPfx(.S1, "?") Then
                .S1 = RmvFstChr(0.1)
            End If
        End If
    End With
Next
SqFldExprAy_RmvFldNmPfxIsQuestionMrk = O
End Function

Private Function SqFldExprAy_RmvTermWithDot(A() As S1S2) As S1S2()
Dim J%, O() As S1S2
O = A
For J = 0 To S1S2_UB(O)
    If HasSubStr(O(J).S1, ".") Then
        O(J).S1 = TakAftRev(O(J).S1, ".")
        O(J).S1 = RmvPfx(O(J).S1, "?")
    End If
Next
SqFldExprAy_RmvTermWithDot = O
End Function

Private Function SqFldExprAy_SetFldNmPfxIsDot(A() As S1S2) As S1S2()
Dim J%, O() As S1S2
O = A
For J = 0 To S1S2_UB(O)
    If FstChr(A(J).S1) = "." Then
        O(J).S1 = RmvFstChr(O(J).S1)
        O(J).S2 = O(J).S1
    End If
Next
SqFldExprAy_SetFldNmPfxIsDot = O
End Function

Private Function SqLnxsAy_StmtsRslt(SqLnxsAy() As Lnx1s, Pm As Dictionary, Sw As Sw) As StmtsRslt
Dim OEr$()
Dim OStmtAy$()
    Dim J%
    For J = 0 To UB(SqLnxsAy)
        With SqLnxs_StmtRslt(SqLnxsAy(J), Pm, Sw)
            Push OStmtAy, .Stmt
            PushAy OEr, .ErLy
        End With
    Next
Dim Z As StmtsRslt
    Z.Stmts = JnDblCrLf(OStmtAy)
    Stop
    '    Z.ErLy = OEr
SqLnxsAy_StmtsRslt = Z
End Function

Private Function SqLin_Evl$(SqLin, EDic As Dictionary)
Dim L$
    L = SqLin
Dim QItm$
Dim IsOpt As Boolean
'    QItm = ParseTerm(L)

Dim Pfx$
    Select Case QItm
    Case "Upd":    Pfx = "Update "
    Case "Set":    Pfx = "|  Set"
    Case "Sel":    Pfx = "Select"
    Case "SelDis": Pfx = "Select Distinct"
    Case "Gp":     Pfx = "|  Group By"
    Case "Fm":     Pfx = "|  From "
    Case "Left":   Pfx = "|  Left Join "
    Case "Jn":     Pfx = "|  Join "
    Case _
        "WhEDic", _
        "WhInStrLis", _
        "WhInNbrLis", _
        "WhBetStr", _
        "WhBetNbr":
                   Pfx = "|  Where "
    Case _
        "AndEDic", _
        "AndInStrLis", _
        "AndInNbrLis", _
        "AndBetStr", _
        "AndBetNbr":
                   Pfx = "|  And "
    Case Else
        Stop
    End Select
Dim Rst$
    Select Case QItm
    Case _
        "Upd", _
        "Fm", _
        "Left", _
        "Jn", _
        "WhEDic", _
        "AndEDic"
                        Rst = L
    Case _
        "Sel", _
        "SelDis"
                        Rst = SqRst_Sel(Rst, EDic)    ' Sq = Sql-Phrase-Itm-Context
    Case "WhInStrLis":  Rst = SqRst_InLis(Rst, EDic, IsStr:=True)
    Case "WhInNbrLis":  Rst = SqRst_InLis(Rst, EDic, IsStr:=False)
    Case "WhBetStr":    Rst = SqRst_Bet(Rst, EDic, IsStr:=True)
    Case "WhBetNbr":    Rst = SqRst_Bet(Rst, EDic, IsStr:=False)
    Case "Gp":          Rst = SqRst_Gp(Rst, EDic)
    Case Else
        Stop
    End Select
SqLin = Pfx & Rst

End Function

Private Function SqLnxs_DrpStmtRslt(SqLnxs As Lnx1s, Ty As e_StmtTy) As StmtRslt1
If Ty = e_DrpStmt Then Exit Function
Dim TnLvs$
With SqLnxs_DrpStmtRslt
    .Stmt = TblNms(TnLvs).DrpStmts
End With
End Function

Private Function SqLnxs_SelStmtRslt(SqLnxs As Lnx1s, Ty As e_StmtTy, Pm As Dictionary, FldSw As Dictionary) As StmtRslt1
If Ty = e_SelStmt Then Exit Function
End Function

Private Function SqLnxs_UpdStmtRslt(SqLnxs As Lnx1s, Ty As e_StmtTy, Pm As Dictionary, FldSw As Dictionary) As StmtRslt1
If Ty <> e_UpdStmt Then Exit Function
End Function

Private Function SqLy_IsSkip(A$(), Ty As e_StmtTy, TblSw As Dictionary) As Boolean
Dim T$: T = SqLy_TarTn(A, Ty)
SqLy_IsSkip = TblSw.Exists(T)
End Function

Private Function SqLy_TarTn$(A$(), Ty As e_StmtTy)

End Function

Private Function SqLy_Ty(Ly$()) As e_StmtTy

End Function

Private Function SqRst_Bet$(A$, EDic As Dictionary, IsStr As Boolean)
Dim ETerm$, T1$, T2$
Dim C$, E$
    Const C1$ = "? Between '?' and '?'"
    Const C2$ = "? Between ? and ?"
    C = IIf(IsStr, C1, C2)
SqRst_Bet = FmtQQ(C, E, T1, T2)
End Function

Private Function SqRst_Gp$(A$, EDic As Dictionary)
Dim ExprAy$(), Ay$()
Stop
'    ExprAy = DicSelIntoSy(EDic, Ay)
SqRst_Gp = SqpGp(ExprAy)
End Function

Private Function SqRst_InLis$(A$, EDic As Dictionary, IsStr As Boolean)
Dim ETerm$, TermAy$()
Ass EDic.Exists(ETerm)
Dim L$, E$
    If IsStr Then L = JnQSngComma(TermAy) Else L = JnComma(TermAy)
SqRst_InLis = FmtQQ("? in ( ? )", E, L)
End Function


Private Function TpEr_Add3(A1 As TpEr, A2 As TpEr, A3 As TpEr) As TpEr
Dim O As TpEr
TpEr_Add3 = O
End Function

Private Function ZZMd() As CodeModule
'Set ZZMd = Md("SqTpSw")
End Function

Private Function ZZPm() As Dictionary
Stop
'Set ZZPm = NewLyDic(ZZPmLy)
End Function

Private Function ZZPmLy() As String()
ZZPmLy = MdResLy(ZZMd, "PmLy")
End Function

Private Sub ZZResPmLy()
'sldkfj skldjf '
' skdfjl
'sdfl sdkfl
End Sub

Private Sub ZZResSqTp()
'-- Rmk: -- is remark
'-- %XX: is prmDicLin
'-- %?XX: is switchPrm, it value must be 0 or 1
'-- ?XX: is switch line
'-- SwitchLin: is ?XXX [OR|AND|EQ|NE] [SwPrm_OR_AND|SwPrm_EQ_NE]
'-- SwPrm_OR_AND: SwTerm ..
'-- SwPrm_EQ_NE:  SwEQ_NE_T1 SwEQ_NE_T2
'-- SwEQ_NE_T1:
'-- SwEQ_NE_T2:
'-- SwTerm:     ?XX|%?XX     -- if %?XX, its value only 1 or 0 is allowed
'-- Only one gp of %XX:
'-- Only one gp of ?XX:
'-- All other gp is sql-statement or sql-statements
'-- sql-statments: Drp xxx xxx
'-- sql-statment: [sel|selDis|upd|into|fm|whBetStr|whBetNbr|whInStrLis|whInNbrLis|andInNbrLis|andInStrLis|gp|jn|left|expr]
'-- optional: Whxxx and Andxxx can have ?-pfx becomes: ?Whxxx and ?Andxxx.  The line will become empty
'==============================================
'Drp Tx TxMbr MbrDta Div Sto Crd Cnt Oup MbrWs
'=============================================
'-- %? means switch, value must be 0 or 1
'%?BrkMbr 0
'%?BrkMbr 0
'%?BrkMbr 0
'%?BrkSto 0
'%?BrkCrd 0
'%?BrkDiv 0
'-- %XXX means txt and optional, allow, blank
'%SumLvl  Y
'%?MbrEmail 1
'%?MbrNm    1
'%?MbrPhone 1
'%?MbrAdr   1
'-- %% mean compulasary
'%%DteFm 20170101
'%%DteTo 20170131
'%LisDiv 1 2
'%LisSto
'%LisCrd
'%CrdExpr ...
'%CrdExpr ...
'%CrdExpr ...
'============================================
'-- EQ & NE t1 only TxtPm is allowed
'--         t2 allow TxtPm, *BLANK, and other text
'?LvlY    EQ %SumLvl Y
'?LvlM    EQ %SumLvl M
'?LvlW    EQ %SumLvl W
'?LvlD    EQ %SumLvl D
'?Y       OR ?LvlD ?LvlW ?LvlM ?LvlY
'?M       OR ?LvlD ?LvlW ?LvlM
'?W       OR ?LvlD ?LvlW
'?D       OR ?LvlD
'?Dte     OR ?LvlD
'?Mbr     OR %?BrkMbr
'?MbrCnt  OR %?BrkMbr
'?Div     OR %?BrkDiv
'?Sto     OR %?BrkSto
'?Crd     OR %?BrkCrd
'?#SEL#Div NE %LisDiv *blank
'?#SEL#Sto NE %LisSto *blank
'?#SEL#Crd NE %LisCrd *blank
'============================================= #Tx
'sel  ?Crd ?Mbr ?Div ?Sto ?Y ?M ?W ?WD ?D ?Dte Amt Qty Cnt
'into #Tx
'fm   SalesHistory
'wh   bet str    %%DteFm %%DteTo
'?and in  strlis Div %LisDiv
'?and in  strlis Sto %LisSto
'?and in  nbrlis Crd %LisCrd
'?gp  ?Crd ?Mbr ?Div ?Sto ?Crd ?Y ?M ?W ?WD ?D ?Dte
'$Crd %CrdExpr
'$Mbr JCMCode
'$Sto
'$Y
'$M
'$W
'$WD
'$D
'$Dte
'$Amt Sum(SHAmount)
'$Qty Sum(SHQty)
'$Cnt Count(SHInvoice+SHSDate+SHRef)
'============================================= #TxMbr
'selDis  Mbr
'fm      #Tx
'into    #TxMbr
'============================================= #MbrDta
'sel   Mbr Age Sex Sts Dist Area
'fm    #TxMbr x
'jn    JCMMember a on x.Mbr = a.JCMMCode
'into  #MbrDta
'$Mbr  x.Mbr
'$Age  DATEDIFF(YEAR,CONVERT(DATETIME ,x.JCMDOB,112),GETDATE())
'$Sex  a.JCMSex
'$Sts  a.JCMStatus
'$Dist a.JCMDist
'$Area a.JCMArea
'==-=========================================== #Div
'?sel Div DivNm DivSeq DivSts
'fm   Division
'into #Div
'?wh in strLis Div %LisDiv
'$Div Dept + Division
'$DivNm LongDies
'$DivSeq Seq
'$DivSts Status
'============================================ #Sto
'?sel Sto StoNm StoCNm
'fm   Location
'into #Sto
'?wh in strLis Loc %LisLoc
'$Sto
'$StoNm
'$StoCNm
'============================================= #Crd
'?sel        Crd CrdNm
'fm          Location
'into        #Crd
'?wh in nbrLis Crd %LisCrd
'$Crd
'$CrdNm
'============================================= #Oup
'sel  ?Crd ?CrdNm ?Mbr ?Age ?Sex ?Sts ?Dist ?Area ?Div ?DivNm ?Sto ?StoNm ?StoCNm ?Y ?M ?W ?WD ?D ?Dte Amt Qty Cnt
'into #Oup
'fm   #Tx x
'left #Crd a on x.Crd = a.Crd
'left #Div b on x.Div = b.Div
'left #Sto c on x.Sto = c.Sto
'left #MbrDta d on x.Mbr = d.Mbr
'wh   JCMCode in (Select Mbr From #TxMbr)
'============================================ #Cnt
'sel ?MbrCnt RecCnt TxCnt Qty Amt
'into #Cnt
'fm  #Tx
'$MbrCnt Count(Distinct Mbr)
'$RecCnt Count(*)
'$TxCnt  Sum(TxCnt)
'$Qty    Sum(Qty)
'$Amt    Sum(Amt)
'============================================
'--
'============================================
'df eror fs--
'============================================
'-- EQ & NE t1 only TxtPm is allowed
'--         t2 allow TxtPm, *BLANK, and other text
'?LvlY    EQ %SumLvl Y
'?LvlM    EQ %SumLvl M
'?LvlW    EQ %SumLvl W
'?LvlD    EQ %SumLvl D
'?Y       OR ?LvlD ?LvlW ?LvlM ?LvlY`
End Sub

Private Sub ZZResSwLy()
'sdfsdf
'sdfsdfa
End Sub

Private Sub ZZResWhereTp()

' wh * Bet * Str|Nbr
' wh * In  * Str|Nbr * Lis
' and * Bet * Str|Nbr
' and * In  * Str|Nbr * Lis
' ==> wh|and * ( ( Bet * Str|Nbr) | (In * Str|Nbr * Lis) )
' ==> wh|and * Bet|In * Str|Nbr * Lis|_
' ==> wh|and * Bet|In * Str|Nbr
End Sub

Private Function ZZSqTp$()
ZZSqTp = MdResStr(Md("SqTp"), "SqTp")
End Function

Private Function ZZSqTpLy() As String()
ZZSqTpLy = SplitCrLf(ZZSqTp)
End Function

Private Function ZZSwBrkAyNoEr() As SwBrk()
'Dim B1 As BrkRslt: B1 = Z1_TpEr(ZZSwBrk)
'Dim B2 As BrkRslt: B2 = Z1_DupNmEr(B1.Rslt)
'ZZSwBrkAyNoEr = B2.Rslt
End Function

Private Function ZZSwLnxAy() As Lnx()
'Dim Ly$(): Ly = MdResLy(Md("SqTpSw"), "SwLy")
'ZZSwLnxAy = LyLnxAy(Ly)
End Function

Private Function ZZSwLy() As String()
ZZSwLy = MdResLy(ZZMd, "SwLy")
End Function

Private Sub FmtSql__Tst()
Dim Tp$: Tp = "Select" & _
"|{?Sel}" & _
"|    {ECrd} Crd," & _
"|    {EAmt} Amt," & _
"|    {EQty} Qty," & _
"|    {ECnt} Cnt," & _
"|  Into #Tx" & _
"|  From SaleHistory" & _
"|  Where SHDate Between '{PFm}' and '{PTo}'" & _
"|    And {EDiv} in ({InDiv})" & _
"|  Group By" & _
"|{?Gp}" & _
"|?M   {ETxM}," & _
"|?W   {ETxW}," & _
"|?D   {ETxD}"
'SR_ = Sales Report
Const ETxWD$ = _
"CASE WHEN TxWD1 = 1 then 'Sun'" & _
"|ELSE WHEN TxWD1 = 2 THEN 'Mon'" & _
"|ELSE WHEN TxWD1 = 3 THEN 'Tue'" & _
"|ELSE WHEN TxWD1 = 4 THEN 'Mon'" & _
"|ELSE WHEN TxWD1 = 5 THEN 'Thu'" & _
"|ELSE WHEN TxWD1 = 6 THEN 'Fri'" & _
"|ELSE WHEN TxWD1 = 7 THEN 'Sat'" & _
"|ELSE Null" & _
"|END END END END END END END"
Dim D As New Dictionary
With D
    .Add "ECrd", "Line-1|Line-2"
    .Add "EAmt", "Sum(SHTxDate)"

End With
Dim Act$: 'Act = FmtSql(Tp, D)
Dim Exp$: Exp = ""
Ass Act = Exp
End Sub

Private Sub LnxAy_SwRslt__Tst()
Dim Act As SwRslt
Stop
'Act = LnxAy_SwRslt(ZZSwLnxAy, ZZPm)
Stop
End Sub

Private Sub ZZ_Validate()
StrBrw Init(ZZSqTp).Validate.ToStr
End Sub

Friend Sub Tst()
ZZ_Validate
End Sub

Attribute VB_Name = "SalRpt"
Option Explicit
Public Const SrpKeyLvs$ = _
"LisDiv    " & _
"LisCrd    " & _
"LisSto    " & _
"BrkDiv    " & _
"BrkSto    " & _
"BrkCrd    " & _
"BrkMbr    " & _
"SumLvl    " & _
"InclNm    " & _
"InclAdr   " & _
"InclPhone " & _
"InclEmail " & _
"FmDte     " & _
"ToDte     "
Type SrPm
'Prm
'    LisDiv . 01 02 03
'    LisCrd . 1 2 3 4
'    LisSto . 001 002 003 004
'    ?BrkDiv . 1
'    ?BrkSto . 1
'    ?BrkCrd . 1
'    ?BrkMbr . 0
'    ?InclNm . 1O.Add "InclNm", True
'    ?InclAdr  . 1
'    ?InclPhone . 1
'    ?InclEmail . 1
'    SumLvl .Y
'    Fm . 20170101
'    To . 20170131
    LisDiv As String
    LisCrd As String
    LisSto As String
    BrkDiv As Boolean
    BrkSto As Boolean
    BrkCrd As Boolean
    BrkMbr As Boolean
    SumLvl As String
    InclNm As Boolean
    InclAdr As Boolean
    InclPhone As Boolean
    InclEmail As Boolean
    FmDte As String
    ToDte As String
    ECrd As String
End Type
Type SrPmx
   ECrd As String
   InDiv As String
   InSto As String
   InCrd As String
   InclFldTxD As Boolean
   InclFldTxY As Boolean
   InclFldTxM As Boolean
   InclFldTxW As Boolean
End Type

Function CrdTyLvs_CrdExpr$(CrdTyLvs$, CrdPfxTyDry())
Const CSub$ = "CrdTyLvs_CrdExpr"
Const CaseWhen$ = "Case When"
Const ElseCaseWhen$ = "|Else Case When"
Dim CrdTyAy%()
    CrdTyAy = AyIntAy(LvsSy(CrdTyLvs))
    Dim StdCrdTyAy%()
    StdCrdTyAy = DrySelDisIntCol(CrdPfxTyDry, 1) ' 1 is colIx which is CrdTyId
    Dim NotExistIdAy%()
        NotExistIdAy = AyMinus(CrdTyAy, StdCrdTyAy)
        If Not AyIsEmp(NotExistIdAy) Then
            Er CSub, "{CrdTyLvs} has item not found in std {CrdTyAy} which is comming from {CrdPfxTyDry}"
        End If
    If AyIsEmp(CrdTyAy) Then CrdTyAy = StdCrdTyAy
Dim NGp%, GpAy$()
    NGp = Sz(CrdTyAy)
    Dim O$(), J%
    For J = 0 To UB(CrdTyAy)
        Push O, CrdTyId_GpItm(CrdTyAy(J), CrdPfxTyDry)
    Next
    GpAy = O

Dim Gp$, ElseN$, EndN$
    ElseN = "|Else " & NGp + 1
    EndN = "|" & StrDup(NGp, "End ")
    Gp = Join(GpAy, ElseCaseWhen)
CrdTyLvs_CrdExpr = CaseWhen & Gp & ElseN & EndN
End Function

Function DftSrpDic() As Dictionary
Dim X As Boolean, Y As New Dictionary
If Not X Then
    X = True
    With Y
        .Add "BrkCrd", False
        .Add "BrkDiv", False
        .Add "BrkMbr", False
        .Add "BrkSto", False
        .Add "LisCrd", ""
        .Add "LisSto", ""
        .Add "LisDiv", ""
        .Add "FmDte", "20170101"
        .Add "ToDte", "20170131"
        .Add "SumLvl", "M"
        .Add "InclAdr", False
        .Add "InclNm", False
        .Add "InclPhone", False
        .Add "InclEmail", False
    End With
End If
Set DftSrpDic = Y
End Function

Function DryOf_CrdPfxTy() As Variant()
Static X As Boolean, Y
If Not X Then
    X = True
    Dim O()
    Push O, Array("134234", 1)
    Push O, Array("12323", 1)
    Push O, Array("2444", 2)
    Push O, Array("2443434", 2)
    Push O, Array("24424", 2)
    Push O, Array("3", 3)
    Push O, Array("5446561", 4)
    Push O, Array("6234341", 5)
    Push O, Array("6234342", 5)
    Y = O
End If
DryOf_CrdPfxTy = Y
End Function

Function SqLoFmtr_OCnt$()

End Function

Function SqLoFmtr_OMbrWsOpt$( _
    BrkMbr As Boolean, _
    InclNm As Boolean, _
    InclAdr As Boolean, _
    InclEmail As Boolean, _
    InclPhone As Boolean)
Const SR_ETMbrWsOptEAge$ = "DateDiff(Year, Convert(DateTime, JCMDOB, 112), GETDATE())"
Const SR_ETMbrWsOptEMbr$ = "JCMCode"
Const ESex$ = "JCMSex"
Const ESts$ = "JCMStatus"
Const EDist$ = "JCMDist"
Const EArea$ = "JCMArea"
Const EAdr$ = "Adr-Express-L1|Adr-Expression-L2"
Const Enm$ = "JCMName"
Const ECNm$ = "JCMCName"
Const EPhone$ = "JCMPhone"
Const EEmail$ = "JCMEmail"

'Sql.X.T.Print MbrDta
'    Sel # Mbr Age Sex Sts Dist Area ?Nm ?Email ?Phone ?Adr
'    Fm  # JCMember
'    Wh  # JCMCode (Select Mbr From #TxMbr)
'Sql.X.T.Print MbrDta.Sel
'    Mbr .JCMCode
'    Age .DateDiff(Year, Convert(DateTime, JCMDOB, 112), GETDATE())
'    Sex .JCMSex
'    Sts .JCMStatus
'    Dist .JCMDist
'    Area .JCMArea
If Not BrkMbr Then Exit Function
Dim Fny$()
    Dim Ay$()
    Ay = LvsSy("Mbr Age Sex Sts Dist Area")
    If InclNm Then Push Ay, "Nm"
    If InclAdr Then Push Ay, "Adr"
    If InclEmail Then Push Ay, "Email"
    If InclPhone Then Push Ay, "Phone"
    Fny = Ay
Dim ExprAy$()
    Dim Dic As New Dictionary
    With Dic
        .Add "Mbr", SR_ETMbrWsOptEMbr
        .Add "Age", SR_ETMbrWsOptEAge
        .Add "Sex", ESex
        .Add "Sts", ESts
        .Add "Dist", EDist
        .Add "Area", EArea
        If InclAdr Then .Add "Adr", EAdr
        If InclNm Then .Add "Nm", Enm
        If InclNm Then .Add "CNm", ECNm
        If InclPhone Then .Add "Phone", EPhone
        If InclEmail Then .Add "Email", EEmail
    End With
    ExprAy = DicSelIntoSy(Dic, Fny)
Dim Sel$, Into$, Fm$, Wh$
    Sel = SqpSel(Fny, ExprAy)
    Into = SqpInto("#MbrDta")
    Fm = SqpFm("JCMember")
    Wh = SqpWh("JCMDCode in (Select Mbr From #TxMbr)")
SqLoFmtr_OMbrWsOpt = Sel & Into & Fm & Wh
End Function

Function SqLoFmtr_OOup$()
Const L$ = _
"Sel " & _
"|Into #Oup" & _
"|Fm #Tx x" & _
"|Left #TxMbr a on x.Mbr = a.JCMMCode"
SqLoFmtr_OOup = RplVBar(L)
End Function

Function SqLoFmtr_TCrd$(BrkCrd As Boolean, InCrd$)
Const FldLvs$ = "Crd CrdNm"
Const ECrd$ = "CrdTyId"
Const ECrdNm$ = "CrdTyNm"
'Sql.X.T.Crd
'    Sel  # Crd Nm
'    Fm   # JR_FrqMbrLis_#CrdTy()
'Sql.X.T.Crd.Sel
'    Crd .CrdTyId
'    Nm .CrdTyNm
If Not BrkCrd Then Exit Function
Dim ExprAy$()
    ExprAy = ApSy(ECrd, ECrdNm)
Dim Sel$, Into$, Fm$, Wh$
    Sel = SqpSelFldLvs(FldLvs, ExprAy)
    Into = SqpInto("#Crd")
    Fm = SqpFm("JR_FrqMbrLis_#CrdTy()")
    Wh = IIf(InCrd = "", "", SqpWh(FmtQQ("? in (?)", ECrd, InCrd)))
SqLoFmtr_TCrd = Sel & Into & Fm & Wh
End Function

Function SqLoFmtr_TDiv$(BrkDiv As Boolean, InDiv$)
If Not BrkDiv Then Exit Function
Const FldLvs$ = "Div DivNm DivSeq DivSts"
Const EDiv$ = "Dept + Division"
Const EDivNm$ = "DivNm"
Const EDivSeq$ = "Seq"
Const EDivSts$ = "Status"
Dim ExprAy$()
    ExprAy = ApSy(EDiv, EDivNm, EDivSeq, EDivSts)

Dim Sel$, Into$, Fm$, Wh$
    Sel = SqpSelFldLvs(FldLvs, ExprAy)
    Into = SqpInto("#Div")
    Fm = SqpFm("Division")
    Wh = "": If InDiv <> "" Then Wh = SqpWh(FmtQQ("? in (?)", EDiv, InDiv))
SqLoFmtr_TDiv = Sel & Into & Fm & Wh
End Function

Function SqLoFmtr_TSto$(BrkSto As Boolean, InSto$)
'Sql.X.T.Sto
'    Sel  # Sto Nm CNm
'    Fm   # LocTbl
'Sql.X.T.Sto.Sel
'    Sto . '0'+Loc_Code
'    Nm .Loc_Name
'    CNm .Loc_CName
If Not BrkSto Then Exit Function
Const ESto$ = "'0'+Loc_Code"
Const EStoNm$ = "Loc_Name"
Const EStoCNm$ = "Loc_CName"
Dim ExprAy$()
    ExprAy = ApSy(ESto, EStoNm, EStoCNm)
Dim Sel$, Into$, Fm$, Wh$
    Sel = SqpSelFldLvs("Sto StoNm StoCNm", ExprAy)
    Into = SqpInto("#Sto")
    Fm = SqpFm("Location")
    Wh = IIf(InSto = "", "", SqpWh(FmtQQ("? in (?)", ESto, InSto)))
SqLoFmtr_TSto = Sel & Into & Fm & Wh
End Function

Function SqLoFmtr_TTx$(P As SrPm)
Dim O$()
Const ECnt$ = "Count(SHInvoice + SHSDate + SHRef)"
Const EAmt$ = "Sum(SHAmount)"
Const EQty$ = "Sum(SHQty)"
Const EMbr$ = "Mbr-Expr"
Const EDiv$ = "Div-Expr"
Const ESto$ = "Sto-Expr"
Const ETxY$ = "SUBSTR(SHSDate,1,4)"
Const ETxM$ = "SUBSTR(SHSDate,5,2)"
Const ETxW$ = "TxW-Expr"
Const ETxD$ = "SUBSTR(SHSDate,7,2)"
Const ETxDte$ = "SUBSTR(SHSDate,1,4)+'/'+SUBSTR(SHSDate,5,2)+'/'+SUBSTR(SHSDate,7,2)"
Dim ECrd$
    ECrd = P.ECrd
Dim Px As SrPmx
    Px = SrPm_SrPmx(P, ECrd)
Dim Fny$()
    Erase O
    Push O, "Crd"
    Push O, "Amt"
    Push O, "Qty"
    Push O, "Cnt"
    If P.BrkMbr Then Push O, "Mbr"
    If P.BrkDiv Then Push O, "Div"
    If P.BrkSto Then Push O, "Sto"
    If Px.InclFldTxY Then Push O, "TxY"
    If Px.InclFldTxM Then Push O, "TxM"
    If Px.InclFldTxW Then Push O, "TxW"
    If Px.InclFldTxD Then Push O, "TxD"
    If Px.InclFldTxD Then Push O, "TxDte"
    Fny = O

Dim ExprAy$()
    Erase O
    Push O, ECrd
    Push O, EAmt
    Push O, EQty
    Push O, ECnt
    If P.BrkMbr Then Push O, EMbr
    If P.BrkDiv Then Push O, EDiv
    If P.BrkSto Then Push O, ESto
    If Px.InclFldTxY Then Push O, ETxY
    If Px.InclFldTxM Then Push O, ETxM
    If Px.InclFldTxW Then Push O, ETxW
    If Px.InclFldTxD Then Push O, ETxD
    If Px.InclFldTxD Then Push O, ETxDte
    ExprAy = O

Dim CrdIn$, DivIn$, StoIn$
    CrdIn = SqpExprIn(ECrd, Px.InCrd)
    DivIn = SqpExprIn(EDiv, Px.InDiv)
    StoIn = SqpExprIn(ESto, Px.InSto)

Dim GpExprAy$()
    Erase O
    Push O, ECrd
    If P.BrkMbr Then Push O, EMbr
    If P.BrkDiv Then Push O, EDiv
    If P.BrkSto Then Push O, ESto
    If Px.InclFldTxY Then Push O, ETxY
    If Px.InclFldTxM Then Push O, ETxM
    If Px.InclFldTxD Then Push O, ETxD
    If Px.InclFldTxD Then Push O, ETxDte
    GpExprAy = O

Dim Sel$, Into$, Fm$, Wh$, AndCrd$, AndSto$, AndDiv$, Gp$
    Sel = SqpSel(Fny, ExprAy)
    Into = SqpInto("#Tx")
    Fm = SqpFm("SaleHistory")
    Wh = SqpWhBetStr("SHDate", P.FmDte, P.ToDte)
    AndCrd = SqpAnd(CrdIn)
    AndDiv = SqpAnd(DivIn)
    AndSto = SqpAnd(StoIn)
    Gp = SqpGp(GpExprAy)
SqLoFmtr_TTx = Sel & Into & Fm & Wh & AndCrd & AndDiv & AndSto & Gp
End Function

Function SqLoFmtr_TTxMbr$(BrkMbr As Boolean)
If Not BrkMbr Then Exit Function
SqLoFmtr_TTxMbr = "Select Distinct Mbr From #Tx Into #TxMbr"
End Function

Function SqLoFmtr_TTxSel$(Fny$(), ExprAy$())
SqLoFmtr_TTxSel = SqpSel(Fny, ExprAy)
End Function

Function SqLoFmtr_TUpdTx$(InclFldTxDte As Boolean)
Const SR_TUpdTxETxWD$ = _
"CASE WHEN TxWD1 = 1 then 'Sun'" & _
"|ELSE WHEN TxWD1 = 2 THEN 'Mon'" & _
"|ELSE WHEN TxWD1 = 3 THEN 'Tue'" & _
"|ELSE WHEN TxWD1 = 4 THEN 'Mon'" & _
"|ELSE WHEN TxWD1 = 5 THEN 'Thu'" & _
"|ELSE WHEN TxWD1 = 6 THEN 'Fri'" & _
"|ELSE WHEN TxWD1 = 7 THEN 'Sat'" & _
"|ELSE Null" & _
"|END END END END END END END"
If Not InclFldTxDte Then Exit Function
SqLoFmtr_TUpdTx = SqpUpd("#Tx") & SqpSet("TxWD", ApSy(SR_TUpdTxETxWD))
End Function

Function SrPm_Dic(P As SrPm) As Dictionary
Dim O As Dictionary
With O
    .Add "", P.BrkCrd
    .Add "", P.BrkDiv
    .Add "", P.BrkMbr
    .Add "", P.BrkSto
    .Add "", P.BrkCrd
    .Add "", P.InclNm
    .Add "", P.InclAdr
    .Add "", P.InclEmail
    .Add "", P.InclPhone
    .Add "", P.LisCrd
    .Add "", P.LisDiv
    .Add "", P.LisSto
    .Add "", P.FmDte
    .Add "", P.ToDte
    .Add "", P.SumLvl
End With
If O.Count <> Sz(LvsSy(SrpKeyLvs)) Then Stop
Set SrPm_Dic = O
End Function

Function SrPm_SrPmx(P As SrPm, ECrd$) As SrPmx
Dim InclFldTxD As Boolean
Dim InclFldTxM As Boolean
Dim InclFldTxW As Boolean
Dim InclFldTxY As Boolean
   Dim SumLvl$
   SumLvl = P.SumLvl
   Select Case SumLvl
   Case "D": InclFldTxD = True
   End Select
   '
   Select Case SumLvl
   Case "D", "W", "M": InclFldTxM = True
   End Select
   '
   Select Case SumLvl
   Case "D", "W": InclFldTxW = True
   End Select
   '
   Select Case SumLvl
   Case "D", "W", "M", "Y": InclFldTxY = True
   End Select

Dim O As SrPmx
With O
   .ECrd = ECrd
   .InclFldTxD = InclFldTxD
   .InclFldTxM = InclFldTxM
   .InclFldTxW = InclFldTxW
   .InclFldTxY = InclFldTxY
   .InCrd = JnComma(LvsSy(P.LisCrd))
   .InDiv = JnComma(AyQuoteSng(LvsSy(P.LisDiv)))
   .InSto = JnComma(AyQuoteSng(LvsSy(P.LisSto)))
End With
SrPm_SrPmx = O
End Function

Function SrPmx_Dic(A As SrPmx) As Dictionary

End Function

Function SrpDic_IsVdt(A As Dictionary) As Boolean
SrpDic_IsVdt = DicHasKeyLvs(A, SrpKeyLvs)
End Function

Function SrpNm_Dic(SrpNm$) As Dictionary
Dim O As Dictionary
Set O = FtDic(SrpNm_Ft(SrpNm))
Ass SrpDic_IsVdt(O)
Set SrpNm_Dic = O
End Function

Sub SrpNm_Dlt(SrpNm$)
Ass SrpNm <> ""
FfnDlt SrpNm_Ft(SrpNm)
End Sub

Sub SrpNm_Dmp(Optional A$)
Dim Ft$: Ft = SrpNm_Ft(A)
Debug.Print "**PrmNm=" & A
Debug.Print "**PrmFt=" & Ft
AyDmp FtLy(Ft)
End Sub

Sub SrpNm_Edt(Optional A$)
FtBrw SrpNm_Ft(A)
End Sub

Sub SrpNm_Ens(Optional A$)
Dim Ft$
Ft = SrpNm_Ft(A)
If FfnIsExist(Ft) Then Exit Sub
AyWrt DicLy(DftSrpDic), Ft
End Sub

Function SrpNm_Ft$(Optional A$)
If A <> "" Then Ass IsNm(A)
SrpNm_Ft = SrpPth & FmtQQ("?.SrPm.txt", A)
End Function

Function SrpNm_Ly(Optional A$) As String()
SrpNm_Ly = FtLy(SrpNm_Ft(A))
End Function

Function SrpNm_Sql$(Optional SrpNm$)
Dim P As SrPm
    P = SrpNm_SrPm(SrpNm)
Dim BrkMbr As Boolean
Dim InclNm As Boolean
Dim InclAdr As Boolean
Dim InclEmail As Boolean
Dim InclPhone As Boolean
    InclNm = P.InclNm
    InclAdr = P.InclAdr
    InclEmail = P.InclEmail
    InclPhone = P.InclPhone
Dim O$(), ECrd$
    ECrd = CrdTyLvs_CrdExpr(P.LisCrd, DryOf_CrdPfxTy)
Push O, SqLoFmtr_Drp
Push O, SqLoFmtr_T(P, ECrd)
Push O, SqLoFmtr_O(BrkMbr, InclNm, InclAdr, InclEmail, InclPhone)
SrpNm_Sql = RplVBar(JnCrLf(O))
End Function

Function SrpNm_SrPm(Optional SrpNm$) As SrPm
Dim D As Dictionary
    Set D = DicLy_Dic(SrpNm_Ly(SrpNm))
Ass SrpDic_IsVdt(D)
Dim O As SrPm
With O
    .BrkCrd = D("BrkCrd")
    .BrkDiv = D("BrkDiv")
    .BrkMbr = D("BrkMbr")
    .BrkSto = D("BrkSto")
    .LisCrd = D("LisCrd")
    .LisSto = D("LisSto")
    .LisDiv = D("LisDiv")
    .FmDte = D("FmDte")
    .ToDte = D("ToDte")
    .SumLvl = D("SumLvl")
    .InclNm = D("InclNm")
    .InclEmail = D("InclNm")
    .InclPhone = D("InclPhone")
    .InclAdr = D("InclAdr")
End With
SrpNm_SrPm = O
End Function

Function SrpNy() As String()
SrpNy = PthFnAy(SrpPth, "*-Prm.txt")
End Function

Function SrpPth$()
SrpPth = TstResPth
End Function

Private Function CrdTyId_GpItm$(CrdTyId%, CrdPfxTyDry())
Dim Ay$(): Ay = CrdTyId_SHMCodeLikAy(CrdTyId, CrdPfxTyDry)
Const Sep$ = " OR"
CrdTyId_GpItm = Join(Ay, Sep) & " THEN " & CrdTyId
End Function

Private Function CrdTyId_SHMCodeLikAy(CrdTyId%, CrdPfxTyDry()) As String()
Dim CrdPfxAy() As String
    Dim Dry(): Dry = DryWh(CrdPfxTyDry, 1, CrdTyId)
    CrdPfxAy = DryStrCol(Dry, 0)

Dim O$(), Pfx
    Dim SHMCodeLik$
    For Each Pfx In CrdPfxAy
        SHMCodeLik = FmtQQ("|SHMCode Like '?%'", Pfx)
        Push O, SHMCodeLik
    Next
O = AyAlignL(O)
CrdTyId_SHMCodeLikAy = O
End Function

Private Function SampleSrPm() As SrPm
Dim O As SrPm
With O
    .LisDiv = "01 02 03"
    .LisCrd = "1 2 3 4"
    .LisSto = "001 002 003"
    .BrkDiv = True
    .BrkSto = True
    .BrkCrd = True
    .BrkMbr = True
    .SumLvl = "Y"
    .FmDte = "20170101"
    .ToDte = "20170131"
    .InclNm = True
    .InclAdr = True
    .InclPhone = True
    .InclEmail = True
End With
SampleSrPm = O
End Function

Private Function SqLoFmtr_Drp$()
SqLoFmtr_Drp = TnLvs_DrpSql("#Tx #TxMbr #MbrDta #Div #Sto #Crd #Cnt #Oup #MbrWs")
End Function

Private Function SqLoFmtr_O$( _
    BrkMbr As Boolean, _
    InclNm As Boolean, _
    InclAdr As Boolean, _
    InclEmail As Boolean, _
    InclPhone As Boolean)
Dim O$()
Push O, SqLoFmtr_OCnt
Push O, SqLoFmtr_OOup
Push O, SqLoFmtr_OMbrWsOpt(BrkMbr, InclNm, InclAdr, InclEmail, InclEmail)
O = AyRmvEmp(O)
SqLoFmtr_O = JnDblCrLf(O)
End Function

Private Function SqLoFmtr_T$(P As SrPm, ECrd$)
Dim O$()
Dim Px As SrPmx
    Px = SrPm_SrPmx(P, ECrd)
With P
Push O, SqLoFmtr_TTx(P)
Push O, SqLoFmtr_TUpdTx(Px.InclFldTxD)
Push O, SqLoFmtr_TTxMbr(.BrkMbr)
Push O, SqLoFmtr_TDiv(.BrkDiv, .LisDiv)
Push O, SqLoFmtr_TSto(.BrkSto, .LisSto)
Push O, SqLoFmtr_TCrd(.BrkCrd, .LisCrd)
End With
SqLoFmtr_T = JnCrLf(O)
End Function

Sub SrpNm_Sql__Tst()
StrBrw SrpNm_Sql
End Sub

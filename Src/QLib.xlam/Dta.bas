Attribute VB_Name = "Dta"
Option Explicit

Sub AssEqDry(A(), B())
If Not DryIsEq(A, B) Then Stop
End Sub

Function AyConst_ValConstDry(A, Constant) As Variant()
If AyIsEmp(A) Then Exit Function
Dim O(), I
For Each I In A
   Push O, Array(I, Constant)
Next
AyConst_ValConstDry = O
End Function

Function AyDt(A, Optional FldNm$ = "Itm", Optional DtNm$ = "Ay") As Dt
Dim O As Dt
O.DtNm = DtNm
O.Fny = ApSy(FldNm)
Dim ODry(), J%
For J = 0 To UB(A)
    Push ODry, Array(A(J))
Next
O.Dry = ODry
AyDt = O
End Function

Function ConstAy_ConstValDry(Cons, A) As Variant()
If AyIsEmp(A) Then Exit Function
Dim O(), I
For Each I In A
   Push O, Array(Cons, I)
Next
ConstAy_ConstValDry = O
End Function

Function DaoTyToSim(T As DataTypeEnum) As eSimTy
Dim O As eSimTy
Select Case T
Case _
   DAO.DataTypeEnum.dbBigInt, _
   DAO.DataTypeEnum.dbByte, _
   DAO.DataTypeEnum.dbCurrency, _
   DAO.DataTypeEnum.dbDecimal, _
   DAO.DataTypeEnum.dbDouble, _
   DAO.DataTypeEnum.dbFloat, _
   DAO.DataTypeEnum.dbInteger, _
   DAO.DataTypeEnum.dbLong, _
   DAO.DataTypeEnum.dbNumeric, _
   DAO.DataTypeEnum.dbSingle
   O = eNbr
Case _
   DAO.DataTypeEnum.dbChar, _
   DAO.DataTypeEnum.dbGUID, _
   DAO.DataTypeEnum.dbMemo, _
   DAO.DataTypeEnum.dbText
   O = eTxt
Case _
   DAO.DataTypeEnum.dbBoolean
   O = eLgc
Case _
   DAO.DataTypeEnum.dbDate, _
   DAO.DataTypeEnum.dbTimeStamp, _
   DAO.DataTypeEnum.dbTime
   O = eDte
Case Else
   O = eOth
End Select
DaoTyToSim = O
End Function

Function DbDs(A As Database, Tny0, Optional DsNm$ = "Ds") As Ds
Dim DtAy1() As Dt
    Dim U%, Tny$()
    Tny = DftNy(Tny0)
    U = UB(Tny)
    ReDim DtAy(U)
    Dim J%
    For J = 0 To U
        DtAy(J) = Dbt(A, Tny(J)).Dt
    Next
Dim O As Ds
    O.DsNm = DsNm
    O.DtAy = DtAy1
DbDs = O
End Function

Function DotNy_Dry(DotNy$()) As Variant()
If AyIsEmp(DotNy) Then Exit Function
Dim O(), I
For Each I In DotNy
   With Brk1(I, ".")
       Push O, ApSy(.S1, .S2)
   End With
Next
DotNy_Dry = O
End Function

Function DrExpLinesCol(Dr, LinesColIx%) As Variant()
Dim A$()
    A = SplitCrLf(CStr(Dr(LinesColIx)))
Dim O()
    Dim IDr
        IDr = Dr
    Dim I
    For Each I In A
        IDr(LinesColIx) = I
        Push O, IDr
    Next
DrExpLinesCol = O
End Function

Sub DrIxAy_Asg(Dr, IxAy%(), ParamArray OAp())
Dim J%
For J = 0 To UB(IxAy)
    If IsObject(OAp(J)) Then
        Set OAp(J) = Dr(IxAy(J))
    Else
        OAp(J) = Dr(IxAy(J))
    End If
Next
End Sub

Function DrLin$(Dr, Wdt%())
Dim UDr%
   UDr = UB(Dr)
Dim O$()
   Dim U1%: U1 = UB(Wdt)
   ReDim O(U1)
   Dim W, V
   Dim J%, V1$
   J = 0
   For Each W In Wdt
       If UDr >= J Then V = Dr(J) Else V = ""
       V1 = AlignL(V, W)
       O(J) = V1
       J = J + 1
   Next
DrLin = Quote(Join(O, " | "), "| * |")
End Function

Sub DrecBrw(A As Drec)
DrecDix(A).Brw
End Sub

Function DrecDix(A As Drec) As Dix
Dim J%, O As New Dictionary
For J = 0 To UB(A.Fny)
   O.Add A.Fny(J), A.Dr(J)
Next
Set DrecDix = Dix(O)
End Function

Sub DrecDmp(A As Drec)
DrecDix(A).Dmp
End Sub

Function DrsAddConstCol(A As Drs, ColNm$, ConstVal) As Drs
Dim O As Drs
    O = A
Push O.Fny, ColNm
O.Dry = DryAddConstCol(O.Dry, ConstVal)
DrsAddConstCol = O
End Function

Function DrsAddRowIxCol(A As Drs) As Drs
Dim O As Drs
    O.Fny = AyIns(A.Fny, "Ix")
Dim ODry()
    If Not AyIsEmp(A.Dry) Then
        Dim J&, Dr
        For Each Dr In A.Dry
            Dr = AyIns(Dr, J): J = J + 1
            Push ODry, Dr
        Next
    End If
O.Dry = ODry
DrsAddRowIxCol = O
End Function

Sub DrsBrw(A As Drs, Optional MaxColWdt& = 100, Optional BrkColNm$, Optional Fnn$)
AyBrw DrsLy(A, MaxColWdt, BrkColNm$), Fnn
End Sub

Function DrsCol(Drs As Drs, ColNm$) As Variant()
Dim ColIx%: ColIx = AyIx(Drs.Fny, ColNm)
DrsCol = DryCol(Drs.Dry, ColIx)
End Function

Sub DrsDmp(Drs As Drs, Optional MaxColWdt& = 100, Optional BrkColNm$)
AyDmp DrsLy(Drs, MaxColWdt, BrkColNm$)
End Sub

Function DrsDrpCol(A As Drs, ColLvs_or_Ny) As Drs
Dim ColNy$(): ColNy = DftNy(ColLvs_or_Ny)
Ass AyHasSubAy(A.Fny, ColNy)
Dim IxAy%()
    IxAy = FnyIxAy(A.Fny, ColNy)
Dim J%
With DrsDrpCol
    .Fny = AyWhExclIxAy(A.Fny, IxAy)
    .Dry = DryRmvColByIxAy(A.Dry, IxAy)
End With
End Function

Function DrsExpLinesCol(Drs As Drs, LinesColNm$) As Drs
Dim Ix%
    Ix = AyIx(Drs.Fny, LinesColNm)
Dim Dry()
    Dim Dr
    For Each Dr In Drs.Dry
        PushAy Dry, DrExpLinesCol(Dr, Ix)
    Next
Dim O As Drs
    O.Fny = Drs.Fny
    O.Dry = Dry
DrsExpLinesCol = O
End Function

Function DrsFldLvs$(A As Drs)
DrsFldLvs = JnSpc(A.Fny)
End Function

Sub DrsLoFmt(A As Drs, At As Range, LoFmtrLy$(), Optional LoNm$)
Dim Lo As ListObject
Set Lo = DrsLo(A, At, LoNm)
'LoFmt Lo, LoFmtrLy
End Sub

Function DrsReOrd(A As Drs, Partial_Fny0) As Drs
Dim ReOrdFny$(): ReOrdFny = DftNy(Partial_Fny0)
Dim IxAy&(): IxAy = AyIxAy(A.Fny, ReOrdFny)
Dim OFny$(): OFny = AyReOrd(A.Fny, IxAy)
Dim ODry(): ODry = DryReOrd(A.Dry, IxAy)
DrsReOrd.Fny = OFny
DrsReOrd.Dry = ODry
End Function

Function DrsRowCnt&(A As Drs, ColNm$, EqVal)
DrsRowCnt = DryRowCnt(A.Dry, AyIx(A.Fny, ColNm), EqVal)
End Function

Function DrsSel(A As Drs, Fny0, Optional CrtEmpColIfReqFldNotFound As Boolean) As Drs
Dim Fny$(): Fny = DftNy(Fny0)
Dim IxAy&()
    If CrtEmpColIfReqFldNotFound Then
        IxAy = AyIxAy(A.Fny, Fny, ChkNotFound:=False)
    Else
        IxAy = AyIxAy(A.Fny, Fny, ChkNotFound:=True)
    End If
Dim O As Drs
    O.Fny = Fny
    O.Dry = DrySel(A.Dry, IxAy, CrtEmpColIfReqFldNotFound)
DrsSel = O
End Function

Function DrsSrt(A As Drs, ColNm$, Optional IsDes As Boolean) As Drs
DrsSrt = NewDrs(A.Fny, DrySrt(A.Dry, AyIx(A.Fny, ColNm), IsDes))
End Function

Function DrsStrCol(Drs As Drs, ColNm$) As String()
DrsStrCol = AySy(DrsCol(Drs, ColNm))
End Function

Function DrsWh(A As Drs, Fld, V) As Drs
Const CSub$ = "DrsWh"
Dim Ix%:
    Ix = AyIx(A.Fny, Fld)
    If Ix = -1 Then Er CSub, "{Fld} is not in {Fny}", Fld, A.Fny
DrsWh.Fny = A.Fny
DrsWh.Dry = DryWh(A.Dry, Ix, V)
End Function

Function DrsWhNotIx(A As Drs, IxAy&()) As Drs
Dim O As Drs
    O.Fny = A.Fny
    Dim J&, I&
    For J = 0 To UB(A.Dry)
        If Not AyHas(IxAy, J) Then
            Push O.Dry, A.Dry(J)
        End If
    Next
DrsWhNotIx = O
End Function

Function DrsWhRow(A As Drs, RowIxAy&()) As Drs
Dim O As Drs
    O.Fny = A.Fny
    Dim J&, I&
    For J = 0 To UB(RowIxAy)
        I = RowIxAy(J)
        Push O.Dry, A.Dry(I)
    Next
DrsWhRow = O
End Function

Function DryAddConstCol(Dry(), ConstVal) As Variant()
If AyIsEmp(Dry) Then Exit Function
Dim N%
   N = Sz(Dry(0))
Dim O()
   Dim Dr, J&
   ReDim O(UB(Dry))
   For Each Dr In Dry
       ReDim Preserve Dr(N)
       Dr(N) = ConstVal
       O(J) = Dr
       J = J + 1
   Next
DryAddConstCol = O
End Function

Sub DryBrw(Dry, Optional MaxColWdt& = 100, Optional BrkColIx% = -1)
AyBrw DryLy(Dry, MaxColWdt, BrkColIx)
End Sub

Function DryCol(Dry, Optional ColIx% = 0) As Variant()
If AyIsEmp(Dry) Then Exit Function
Dim O(), Dr
For Each Dr In Dry
   Push O, Dr(ColIx)
Next
DryCol = O
End Function

Function DryColSet(Dry(), Col_Ix%) As Dictionary
Dim O As New Dictionary
If Not AyIsEmp(Dry) Then
    Dim Dr
    For Each Dr In Dry
        SetPush O, Dr(Col_Ix)
    Next
End If
Set DryColSet = O
End Function

Function DryCvCellToStr(Dry, ShwZer As Boolean) As Variant()
Dim O(), Dr
For Each Dr In Dry
   Push O, DrValCellStr(Dr, ShwZer)
Next
DryCvCellToStr = O
End Function

Sub DryDmp(Dry)
AyDmp DryLy(Dry)
End Sub

Function DryDrIx_IsBrk(Dry, DrIx&, BrkColIx%) As Boolean
If AyIsEmp(Dry) Then Exit Function
If DrIx = 0 Then Exit Function
If DrIx = UB(Dry) Then Exit Function
If Dry(DrIx)(BrkColIx) = Dry(DrIx - 1)(BrkColIx) Then Exit Function
DryDrIx_IsBrk = True
End Function

Function DryIntCol(Dry(), ColIx%) As Integer()
DryIntCol = AyIntAy(DryCol(Dry, ColIx))
End Function

Function DryIsEq(A(), B()) As Boolean
Dim N&: N = Sz(A)
If N <> Sz(B) Then Exit Function
If N = 0 Then DryIsEq = True: Exit Function
Dim J&, Dr
For Each Dr In A
   If Not AyIsEq(Dr, B(J)) Then Exit Function
   J = J + 1
Next
DryIsEq = True
End Function

Function DryKeyGpAy(Dry(), K_Ix%, Gp_Ix%) As Variant()
If AyIsEmp(Dry) Then Exit Function
Dim J%, O, K, GpAy(), O_Ix&, Gp, Dr, K_Ay()
For Each Dr In Dry
    K = Dr(K_Ix)
    Gp = Dr(Gp_Ix)
    O_Ix = AyIx(K_Ay, K)
    If O_Ix = -1 Then
        Push K_Ay, K
        Push O, Array(K, Array(Gp))
    Else
        Push O(O_Ix)(1), Gp
    End If
Next
DryKeyGpAy = O
End Function

Function DryLy(A, Optional MaxColWdt& = 100, Optional BrkColIx% = -1, Optional ShwZer As Boolean) As String()
If VarIsEmp(A) Then Exit Function
Dim A1()
    A1 = DryCvCellToStr(A, ShwZer)
Dim Hdr$
    Dim W%(): W = DryWdtAy(A1, MaxColWdt)
    If AyIsEmp(W) Then Exit Function
    Dim HdrAy$()
    ReDim HdrAy(UB(W))
    Dim J%
    For J = 0 To UB(W)
        HdrAy(J) = String(W(J), "-")
    Next
    Hdr = Quote(Join(HdrAy, "-|-"), "|-*-|")

Dim O$()
    Dim Dr, DrIx&, IsBrk As Boolean
    Push O, Hdr
    If BrkColIx >= 0 Then
        For Each Dr In A1
            IsBrk = DryDrIx_IsBrk(A, DrIx, BrkColIx)
            If IsBrk Then Push O, Hdr
            Push O, DrLin(Dr, W)
            DrIx = DrIx + 1
        Next
    Else
        For Each Dr In A1
            Push O, DrLin(Dr, W)
        Next
    End If
    Push O, Hdr
DryLy = O
End Function

Function DryLy_InsBrkLin(DryLy$(), ColIx%) As String()
If Sz(DryLy) = 2 Then DryLy_InsBrkLin = DryLy: Exit Function
Dim Hdr$: Hdr = DryLy(0)
Dim Fm&, L%
   Dim N%: N = ColIx + 1
   Dim P1&, P2&
   P1 = InstrN(Hdr, "|", N)
   P2 = InStr(P1 + 1, Hdr, "|")
   Fm = P1 + 1
   L = P2 - P1 - 1
Dim O$()
   Push O, DryLy(0)
   Dim LasV$: LasV = Mid(DryLy(1), Fm, L)
   Dim J&
   Dim V$
   For J = 1 To UB(DryLy) - 1
       V = Mid(DryLy(J), Fm, L)
       If LasV <> V Then
           Push O, Hdr
           LasV = V
       End If
       Push O, DryLy(J)
   Next
   Push O, AyLasEle(DryLy)
DryLy_InsBrkLin = O
End Function

Function DryMge(Dry, MgeIx%, Sep$) As Variant()
Dim O(), J%
Dim Ix%
For J = 0 To UB(Dry)
   Ix = DryMgeIx(O, Dry(J), MgeIx)
   If Ix = -1 Then
       Push O, Dry(J)
   Else
       O(Ix)(MgeIx) = O(Ix)(MgeIx) & Sep & Dry(J)(MgeIx)
   End If
Next
DryMge = O
End Function

Function DryMgeIx&(Dry, Dr, MgeIx%)
Dim O&, D, J%
For O = 0 To UB(Dry)
   D = Dry(O)
   For J = 0 To UB(Dr)
       If J <> MgeIx Then
           If Dr(J) <> D(J) Then GoTo Nxt
       End If
   Next
   DryMgeIx = O
   Exit Function
Nxt:
Next
DryMgeIx = -1
End Function

Function DryNCol%(Dry)
Dim Dr, O%, M%
For Each Dr In Dry
   M = Sz(Dr)
   If M > O Then O = M
Next
DryNCol = O
End Function

Function DryReOrd(Dry, PartialIxAy&()) As Variant()
If AyIsEmp(Dry) Then Exit Function
Dim Dr, O()
For Each Dr In Dry
   Push O, AyReOrd(Dr, PartialIxAy)
Next
DryReOrd = O
End Function

Function DryRmvColByIxAy(Dry, IxAy%()) As Variant()
If AyIsEmp(Dry) Then Exit Function
Dim O(), Dr
For Each Dr In Dry
   Push O, AyWhExclIxAy(Dr, IxAy)
Next
DryRmvColByIxAy = O
End Function

Function DryRowCnt&(Dry, ColIx&, EqVal)
If AyIsEmp(Dry) Then Exit Function
Dim J&, O&, Dr
For Each Dr In Dry
   If Dr(ColIx) = EqVal Then O = O + 1
Next
DryRowCnt = O
End Function

Function DrySel(A(), ColIxAy&(), Optional CrtEmpColIfReqFldNotFound As Boolean) As Variant()
Dim O(), Dr
If AyIsEmp(A) Then Exit Function
For Each Dr In A
   Push O, AyWhIxAy(Dr, ColIxAy, CrtEmpColIfReqFldNotFound)
Next
DrySel = O
End Function

Function DrySelDis(Dry(), ColIx%) As Variant()
If AyIsEmp(Dry) Then Exit Function
Dim Dr, O()
For Each Dr In Dry
   PushNoDup O, Dr(ColIx)
Next
DrySelDis = O
End Function

Function DrySelDisIntCol(Dry(), ColIx%) As Integer()
DrySelDisIntCol = AyIntAy(DrySelDis(Dry, ColIx))
End Function

Function DrySq(Dry, Optional NColOpt% = 0) As Variant()
If AyIsEmp(Dry) Then Exit Function
Dim NRow&, NCol&
   If NColOpt <= 0 Then NCol = DryNCol(Dry)
   NRow = Sz(Dry)
Dim O()
   ReDim O(1 To NRow, 1 To NCol)
Dim C%, R&, Dr
   R = 0
   For Each Dr In Dry
       R = R + 1
       For C = 0 To UB(Dr)
           O(R, C + 1) = Dr(C)
       Next
   Next
DrySq = O
End Function

Function DrySrt(Dry, ColIx%, Optional IsDes As Boolean) As Variant()
Dim Col: Col = DryCol(Dry, ColIx)
Dim Ix&(): Ix = AySrtInToIxAy(Col, IsDes)
Dim J%, O()
For J = 0 To UB(Ix)
   Push O, Dry(Ix(J))
Next
DrySrt = O
End Function

Function DryStrCol(Dry, Optional ColIx% = 0) As String()
DryStrCol = AySy(DryCol(Dry, ColIx))
End Function

Function DryWdtAy(Dry, Optional MaxColWdt& = 100) As Integer()
Const CSub$ = "DryWdtAy"
If AyIsEmp(Dry) Then Exit Function
Dim O%()
   Dim Dr, UDr%, U%, V, L%, J%
   U = -1
   For Each Dr In Dry
       If Not VarIsStrAy(Dr) Then Er CSub, "This routine should call DryCvFmtEachCell first so that each cell is ValCellStr as a string.|Now some Dr in given-Dry is not a StrAy, but[" & TypeName(Dr) & "]"
       UDr = UB(Dr)
       If UDr > U Then ReDim Preserve O(UDr): U = UDr
       If AyIsEmp(Dr) Then GoTo Nxt
       For J = 0 To UDr
           V = Dr(J)
           L = Len(V)

           If L > O(J) Then O(J) = L
       Next
Nxt:
   Next
Dim M%
M = MaxColWdt
For J = 0 To UB(O)
   If O(J) > M Then O(J) = M
Next
DryWdtAy = O
End Function

Function DryWh(Dry(), ColIx%, EqVal) As Variant()
Dim O()
Dim J&
For J = 0 To UB(Dry)
   If Dry(J)(ColIx) = EqVal Then Push O, Dry(J)
Next
DryWh = O
End Function

Function DryWs(Dry, Optional WsNm$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = NewWs(WsNm, Vis:=True)
DryRg Dry, WsA1(O)
Set DryWs = O
End Function

Sub DsAddDt(ODs As Ds, T As Dt)
If DsHasDt(ODs, T.DtNm) Then Err.Raise 1, , FmtQQ("DsAddDt: Ds[?] already has Dt[?]", ODs.DsNm, T.DtNm)
Dim N%: N = DtAySz(ODs.DtAy)
ReDim Preserve ODs.DtAy(N)
ODs.DtAy(N) = T
End Sub

Sub DsBrw(A As Ds)
AyBrw DsLy(A)
End Sub

Sub DsDmp(A As Ds)
AyDmp DsLy(A)
End Sub

Function DsHasDt(A As Ds, DtNm) As Boolean
If DsIsEmp(A) Then Exit Function
Dim J%
For J = 0 To UBound(A.DtAy)
    If A.DtAy(J).DtNm = DtNm Then DsHasDt = True: Exit Function
Next
End Function

Function DsIsEmp(A As Ds) As Boolean
DsIsEmp = DtAy_IsEmp(A.DtAy)
End Function

Function DsLy(A As Ds, Optional MaxColWdt& = 1000, Optional DtBrkLinMapStr$) As String()
Dim O$()
    Push O, "*Ds " & A.DsNm & "=================================================="
Dim Dic As Dictionary ' DicOf_Tn_to_BrkColNm
    Set Dic = MapStr_Dic(DtBrkLinMapStr)
If Not DtAy_IsEmp(A.DtAy) Then
    Dim J%, DtNm$, Dt As Dt, BrkColNm$
    For J = 0 To UBound(A.DtAy)
        Dt = A.DtAy(J)
        DtNm$ = Dt.DtNm
        If Dic.Exists(DtNm) Then BrkColNm = Dic(DtNm) Else BrkColNm = ""
        PushAy O, DtLy(Dt, MaxColWdt, BrkColNm)
    Next
End If
DsLy = O
End Function

Function DtAySz%(DtAy() As Dt)
On Error Resume Next
DtAySz = UBound(DtAy) + 1
End Function

Function DtAy_IsEmp(A() As Dt) As Boolean
DtAy_IsEmp = DtAySz(A) = 0
End Function

Sub DtBrw(Dt As Dt, Optional Fnn)
AyBrw DtLy(Dt), IIf(VarIsEmp(Fnn), Dt.DtNm, Fnn)
End Sub

Function DtCsvLy(A As Dt) As String()
Dim O$()
Dim QQStr$
Dim Dr
Push O, JnComma(AyQuoteDbl(A.Fny))
For Each Dr In A.Dry
   Push O, FmtQQAv(QQStr, Dr)
Next
End Function

Sub DtDmp(A As Dt)
AyDmp DtLy(A)
End Sub

Function DtDrpCol(A As Dt, ColLvs_or_Ny) As Dt
Dim B As Drs: B = DtDrs(A)
Dim C As Drs: C = DrsDrpCol(B, ColLvs_or_Ny)
DtDrpCol = NewDt(A.DtNm, C.Fny, C.Dry)
End Function

Function DtDrs(A As Dt) As Drs
Dim O As Drs
O.Fny = A.Fny
O.Dry = A.Dry
DtDrs = O
End Function

Function DtIsEmp(A As Dt) As Boolean
DtIsEmp = AyIsEmp(A.Dry)
End Function

Function DtLy(A As Dt, Optional MaxColWdt& = 100, Optional BrkColNm$, Optional ShwZer As Boolean) As String()
Dim O$()
   Push O, "*Tbl " & A.DtNm
   PushAy O, DrsLy(DtDrs(A), MaxColWdt, BrkColNm, ShwZer)
DtLy = O
End Function

Function DtNmDrs_Dt(A$, Drs As Drs) As Dt
With DtNmDrs_Dt
    .DtNm = A
    .Fny = Drs.Fny
    .Dry = Drs.Dry
End With
End Function

Function DtReOrd(A As Dt, ColLvs$) As Dt
Dim ReOrdFny$(): ReOrdFny = LvsSy(ColLvs)
Dim IxAy&(): IxAy = AyIxAy(A.Fny, ReOrdFny)
Dim OFny$(): OFny = AyReOrd(A.Fny, IxAy)
Dim ODry(): ODry = DryReOrd(A.Dry, IxAy)
DtReOrd.DtNm = A.DtNm
DtReOrd.Fny = OFny
DtReOrd.Dry = ODry
End Function

Function DtSrt(A As Dt, ColNm$, Optional IsDes As Boolean) As Dt
DtSrt = NewDtByDrs(A.DtNm, DrsSrt(DtDrs(A), ColNm, IsDes))
End Function

Sub Fiy(Fny$(), FldLvs$, ParamArray OAp())
'Fiy=Field Index Array
Dim A$(): A = SplitSpc(FldLvs)
Dim I&(): I = AyIxAy(Fny, A)
Dim J%
For J = 0 To UB(I)
    OAp(J) = I(J)
Next
End Sub

Function IsSimTyLvs(A$) As Boolean
Dim Ay$(): Ay = LvsSy(A)
If AyIsEmp(Ay) Then Exit Function
Dim I
For Each I In Ay
   If Not IsSimTyStr(Ay) Then Exit Function
Next
IsSimTyLvs = True
End Function

Function IsSimTyStr(S) As Boolean
Select Case UCase(S)
Case "TXT", "NBR", "LGC", "DTE", "OTH": IsSimTyStr = True
End Select
End Function

Function ItrCntByBoolPrp&(A, BoolPrpNm$)
If A.Count = 0 Then Exit Function
Dim O, Cnt&
For Each O In A
    If CallByName(O, BoolPrpNm, VbGet) Then
        Cnt = Cnt + 1
    End If
Next
ItrCntByBoolPrp = Cnt
End Function

Function ItrDrs(Itr, PrpNy0) As Drs
Dim Ny$()
    Ny = DftNy(PrpNy0)
Dim Dry()
    Dim Obj
    If Itr.Count > 0 Then
        For Each Obj In Itr
            Push Dry, ObjPrpDr(Obj, Ny)
        Next
    End If
Dim O As Drs
    O.Fny = Ny
    O.Dry = Dry
ItrDrs = O
End Function

Function ItrItmByPrp(A, PrpNm$, PrpV)
Dim O, V
If A.Count > 0 Then
    For Each O In A
        V = CallByName(O, PrpNm, VbGet)
        If V = PrpV Then
            Asg O, ItrItmByPrp
            Exit Function
        End If
    Next
End If
End Function

Function ItrNy(A, Optional Lik$ = "*") As String()
Dim O$(), Obj, N$
If A.Count > 0 Then
    For Each Obj In A
        N = Obj.Name
        If N Like Lik Then Push O, N
    Next
End If
ItrNy = O
End Function

Function NewDDLines(Ly$()) As DDLines
Dim O As New DDLines
Set NewDDLines = O.Init(Ly)
End Function

Function NewDrByLvs(Lvs$, TyAy() As eSimTy) As Variant()

End Function

Function NewDrs(Fny0, Dry) As Drs
NewDrs.Fny = DftNy(Fny0)
NewDrs.Dry = Dry
End Function

Function NewDrsByLines(DrsLines$) As Drs
NewDrsByLines = NewDrsByLy(SplitCrLf(DrsLines))
End Function

Function NewDrsByLy(DrsLy$()) As Drs
Dim Fny$(): Fny = LvsSy(DrsLy(0))
Dim J&, Dry()
If IsSimTyLvs(DrsLy(2)) Then
    Dim TyAy() As eSimTy
    For J = 3 To UB(DrsLy)
        Push Dry, NewDrByLvs(DrsLy(J), TyAy)
    Next
Else
    For J = 2 To UB(DrsLy)
        Push Dry, LvsSy(DrsLy(J))
    Next
End If
NewDrsByLy = NewDrs(Fny, Dry)
End Function

Function NewDrsByVbl(DrsVbl$) As Drs
'SpecStr:Vbl:VbarLine
'SpecStr:DrsVbl:Data-record-set-vbar-line
NewDrsByVbl = NewDrsByLy(SplitVBar(DrsVbl))
End Function

Function NewDs(A() As Dt, Optional DsNm$ = "Ds") As Ds
NewDs.DsNm = DsNm
NewDs.DtAy = A
End Function

Function NewDt(DtNm$, Fny0, Dry) As Dt
NewDt.Dry = Dry
NewDt.Fny = DftNy(Fny0)
NewDt.DtNm = DtNm
End Function

Function NewDtByDrs(DtNm$, A As Drs) As Dt
NewDtByDrs = NewDt(DtNm, A.Fny, A.Dry)
End Function

Function NewSimTy(SimTyStr$) As eSimTy
Dim O As eSimTy
Select Case UCase(SimTyStr)
Case "TXT": O = eTxt
Case "NBR": O = eNbr
Case "LGC": O = eLgc
Case "DTE": O = eDte
Case Else: O = eOth
End Select
NewSimTy = O
End Function

Function ObjPrpDr(Obj, PrpNy0) As Variant()
Dim Ny$(): Ny = DftNy(PrpNy0)
Dim U%
    U = UB(Ny)
Dim O()
    ReDim O(U)
    Dim J%
    For J = 0 To U
        O(J) = CallByName(Obj, Ny(J), VbGet)
    Next
ObjPrpDr = O
End Function

Function S1S2Ay_Drs(A() As S1S2) As Drs
S1S2Ay_Drs.Fny = SplitSpc("S1 S2")
S1S2Ay_Drs.Dry = S1S2Ay_Dry(A)
End Function

Function S1S2Ay_Dry(A() As S1S2) As Variant()
Dim O()
Dim J%
For J = 0 To S1S2_UB(A)
   With A(J)
       Push O, Array(.S1, .S2)
   End With
Next
S1S2Ay_Dry = O
End Function

Function SampleDt() As Dt
Dim O As Dt
O.DtNm = "Sample"
O.Dry = Array(Array(1))
O.Fny = LvsSy("A B C")
End Function

Sub SetPush(A As Dictionary, K)
If A.Exists(K) Then Exit Sub
A.Add K, Empty
End Sub

Function SimTy_QuoteTp$(A As eSimTy)
Const CSub$ = "SimTyQuoteTp"
Dim O$
Select Case A
Case eTxt: O = "'?'"
Case eNbr, eLgc: O = "?"
Case eDte: O = "#?#"
Case Else
   Er CSub, "Given {eSimTy} should be [eTxt eNbr eDte eLgc]", A
End Select
SimTy_QuoteTp = O
End Function

Function SqNCol%(A)
On Error Resume Next
SqNCol = UBound(A, 2)
End Function

Function SqNRow%(A)
On Error Resume Next
SqNRow = UBound(A, 1)
End Function

Function TitAy_Sq(TitAy$())
Dim UFld%: UFld = UB(TitAy)
Dim ColVBar()
    ReDim ColVBar(UFld)
    Dim J%
    For J = 0 To UFld
        ColVBar(J) = AyTrim(SplitVBar(TitAy(J)))
    Next
Dim NRow%
    Dim M%, VBar$()
    For J = 0 To UB(ColVBar)
        VBar = ColVBar(J)
        M = Sz(VBar)
        If M > NRow Then NRow = M
    Next
Dim O()
    Dim I%
    ReDim O(1 To NRow, 1 To UFld + 1)
    For J = 0 To UFld
        VBar = ColVBar(J)
        For I = 0 To UB(VBar)
            O(I + 1, J + 1) = VBar(I)
        Next
    Next
TitAy_Sq = O
End Function

Function VblLy_Dry(A$()) As Variant()
If AyIsEmp(A) Then Exit Function
Dim O()
   Dim I
   For Each I In A
       Push O, SyTrim(SplitVBar(CStr(I)))
   Next
VblLy_Dry = O
End Function

Private Sub DbDs__Tst()
Dim Ds As Ds
Ds = DbDs(CurDb, "Permit PermitD")
Stop
End Sub

Private Sub DrsSel__Tst()
'DrsBrw DrsSel(Vmd.MthDrs, "MthNm Mdy Ty MdNm")
'DrsBrw Vmd.MthDrs
End Sub

Private Sub DsWb__Tst()
Dim Wb As Workbook
Set Wb = DsWb(DbDs(CurDb, "Permit PermitD"))
WbVis Wb
Stop
Wb.Close False
End Sub

Sub ItrDrs__Tst()
DrsBrw ItrDrs(Dbt(SampleDb_DutyPrepare, "Permit").Flds, "Name Type Required")
'DrsBrw ItrDrs(Application.VBE.VBProjects, "Name Type")
End Sub

Private Sub TitAy_Sq__Tst()
Dim A$()
Push A, "ksdf | skdfj  |skldf jf"
Push A, "skldf|sdkfl|lskdf|slkdfj"
Push A, "askdfj|sldkf"
Push A, "fskldf"
SqBrw TitAy_Sq(A)
End Sub

Sub VblLy_Dry__Tst()
Dim VblLy$()
Dim Exp$()
Push VblLy, "|lskdf|sdlf|lsdkf"
Push VblLy, "|lsdf|"
Push VblLy, "|lskdfj|sdlfk|sdlkfj|sdklf|skldf|"
Push VblLy, "|sdf"
Dim Act()
Act = VblLy_Dry(VblLy)
DryBrw Act
End Sub

Attribute VB_Name = "M_Dry"
Option Explicit

Property Get Ay_Dry_By_AddConst_Aft_Ay(A, Constant) As Variant()
If Sz(A) = 0 Then Exit Property
Dim O(), I
For Each I In A
   Push O, Array(I, Constant)
Next
Ay_Dry_By_AddConst_Aft_Ay = O
End Property

Property Get Ay_Dry_By_AddConst_Bef_Ay(A, Constant) As Variant()
If Sz(A) = 0 Then Exit Property
Dim O(), I
For Each I In A
   Push O, Array(Constant, I)
Next
Ay_Dry_By_AddConst_Bef_Ay = O
End Property

Property Get DaoTy_SimTy(T As DataTypeEnum) As eSimTy
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
DaoTy_SimTy = O
End Property

Property Get DotNy_Dry(DotNy$()) As Variant()
If AyIsEmp(DotNy) Then Exit Property
Dim O(), I
For Each I In DotNy
   With Brk1(I, ".")
       Push O, ApSy(.S1, .S2)
   End With
Next
DotNy_Dry = O
End Property

Property Get DrExpLinesCol(Dr, LinesColIx%) As Variant()
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
End Property

Property Get DrLin$(Dr, Wdt%())
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
End Property

Property Get DrsDt(A As Drs, Dt$, Drs As Drs) As Dt
Set DrsDt = Drs()
End Property

Property Get DrsExpLinesCol(A As Drs, LinesColNm$) As Drs
Dim Ix%
    Ix = AyIx(A.Fny, LinesColNm)
Dim Dry()
    Dim Dr
    For Each Dr In A.Dry
        PushAy Dry, DrExpLinesCol(Dr, Ix)
    Next
Set DrsExpLinesCol = Drs(A.Fny, Dry)
End Property

Property Get DrsLy_Drs(DrsLy$()) As Drs
Dim Fny$(): Fny = SslSy(DrsLy(0))
Dim J&, Dry()
If IsSimTySsl(DrsLy(2)) Then
    Dim TyAy() As eSimTy
    For J = 3 To UB(DrsLy)
        Push Dry, SslDr(DrsLy(J), TyAy)
    Next
Else
    For J = 2 To UB(DrsLy)
        Push Dry, SslSy(DrsLy(J))
    Next
End If
Set DrsLy_Drs = Drs(Fny, Dry)
End Property

Property Get DrsRowCnt&(A As Drs, ColNm$, EqVal)
DrsRowCnt = DryRowCnt(A.Dry, AyIx(A.Fny, ColNm), EqVal)
End Property

Property Get DrsSel(A As Drs, Fny0, Optional CrtEmpColIfReqFldNotFound As Boolean) As Drs
Dim Fny$(): Fny = DftNy(Fny0)
Dim IxAy&()
    If CrtEmpColIfReqFldNotFound Then
        IxAy = AyIxAy(A.Fny, Fny, ChkNotFound:=False)
    Else
        IxAy = AyIxAy(A.Fny, Fny, ChkNotFound:=True)
    End If
Dim Dry()
    Dry = DrySel(A.Dry, IxAy, CrtEmpColIfReqFldNotFound)
DrsSel = Drs(Fny, Dry)
End Property

Property Get DrsSrt(A As Drs, ColNm$, Optional IsDes As Boolean) As Drs
Set DrsSrt = Drs(A.Fny, DrySrt(A.Dry, AyIx(A.Fny, ColNm), IsDes))
End Property

Property Get DrsVbl_Drs(DrsVbl$) As Drs
'SpecStr:Vbl:VbarLine
'SpecStr:DrsVbl:Data-record-set-vbar-line
DrsVbl_Drs = DrsLy_Drs(SplitVBar(DrsVbl))
End Property

Property Get DrsWh(A As Drs, Fld, V) As Drs
Const CSub$ = "DrsWh"
Dim Ix%:
    Ix = AyIx(A.Fny, Fld)
    If Ix = -1 Then Er CSub, "{Fld} is not in {Fny}", Fld, A.Fny
Set DrsWh = Drs(A.Fny, DryWh(A.Dry, Ix, V))
End Property

Property Get DrsWhNotRowIxAy(A As Drs, NotRowIxAy&()) As Drs
Dim O()
    Dim Dry()
    Dry = A.Dry
    Dim I, Ix&()
    Ix = AyMinus(UIxAy(UB(Dry)), NotRowIxAy)
    For Each I In Ix
        Push O, Dry(I)
    Next
Set DrsWhNotRowIxAy = Drs(A.Fny, O)
End Property

Property Get DrsWhRowIxAy(A As Drs, RowIxAy&()) As Drs
Dim O()
    Dim J&, I, Dry()
    Dry = A.Dry
    For Each I In RowIxAy
        Push O, Dry(I)
    Next
Set DrsWhRowIxAy = Drs(A.Fny, O)
End Property

Property Get DryAddConstCol(Dry(), ConstVal) As Variant()
If AyIsEmp(Dry) Then Exit Property
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
End Property

Property Get DryCol(Dry, Optional ColIx% = 0) As Variant()
If AyIsEmp(Dry) Then Exit Property
Dim O(), Dr
For Each Dr In Dry
   Push O, Dr(ColIx)
Next
DryCol = O
End Property

Property Get DryCol_Into(A, ColIx%, OIntoAy)
Dim O: O = OIntoAy: Erase O
If Sz(A) = 0 Then
    DryCol_Into = O
    Exit Property
End If
Dim Dr, J&
ReDim O(UB(A))
For Each Dr In A
    If UB(Dr) >= ColIx Then
        O(J) = Dr(ColIx)
    End If
    J = J + 1
Next
End Property

Property Get DryCvCellToStr(Dry, ShwZer As Boolean) As Variant()
Dim O(), Dr
For Each Dr In Dry
   Push O, AyCellSy(Dr, ShwZer)
Next
DryCvCellToStr = O
End Property

Property Get DryDrIx_IsBrk(Dry, DrIx&, BrkColIx%) As Boolean
If AyIsEmp(Dry) Then Exit Property
If DrIx = 0 Then Exit Property
If DrIx = UB(Dry) Then Exit Property
If Dry(DrIx)(BrkColIx) = Dry(DrIx - 1)(BrkColIx) Then Exit Property
DryDrIx_IsBrk = True
End Property

Property Get DryIntCol(A, ColIx%) As Integer()
DryIntCol = DryCol_Into(A, ColIx, EmpIntAy)
End Property

Property Get DryIsEq(A(), B()) As Boolean
Dim N&: N = Sz(A)
If N <> Sz(B) Then Exit Property
If N = 0 Then DryIsEq = True: Exit Property
Dim J&, Dr
For Each Dr In A
   If Not AyIsEq(Dr, B(J)) Then Exit Property
   J = J + 1
Next
DryIsEq = True
End Property

Property Get DryKeyGpAy(Dry(), K_Ix%, Gp_Ix%) As Variant()
If AyIsEmp(Dry) Then Exit Property
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
End Property

Property Get DryKix_GpAy(A, Kix%, Gix%) As Variant()
If Sz(A) = 0 Then Exit Property
Dim J%, O, K, GpAy(), O_Ix&, Gp, Dr, K_Ay()
For Each Dr In A
    K = Dr(Kix)
    Gp = Dr(Gix)
    O_Ix = AyIx(K_Ay, K)
    If O_Ix = -1 Then
        Push K_Ay, K
        Push O, Array(K, Array(Gp))
    Else
        Push O(O_Ix)(1), Gp
    End If
Next
DryKix_GpAy = O
End Property

Property Get DryLy(A, Optional MaxColWdt& = 100, Optional BrkColIx% = -1, Optional ShwZer As Boolean) As String()
If IsEmp(A) Then Exit Property
Dim A1()
    A1 = DryCvCellToStr(A, ShwZer)
Dim Hdr$
    Dim W%(): W = DryWdtAy(A1, MaxColWdt)
    If AyIsEmp(W) Then Exit Property
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
End Property

Property Get DryLy_InsBrkLin(DryLy$(), ColIx%) As String()
If Sz(DryLy) = 2 Then DryLy_InsBrkLin = DryLy: Exit Property
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
End Property

Property Get DryMge(Dry, MgeIx%, Sep$) As Variant()
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
End Property

Property Get DryMgeIx&(Dry, Dr, MgeIx%)
Dim O&, D, J%
For O = 0 To UB(Dry)
   D = Dry(O)
   For J = 0 To UB(Dr)
       If J <> MgeIx Then
           If Dr(J) <> D(J) Then GoTo Nxt
       End If
   Next
   DryMgeIx = O
   Exit Property
Nxt:
Next
DryMgeIx = -1
End Property

Property Get DryNCol%(Dry)
Dim Dr, O%, M%
For Each Dr In Dry
   M = Sz(Dr)
   If M > O Then O = M
Next
DryNCol = O
End Property

Property Get DryReOrd(Dry, PartialIxAy&()) As Variant()
If AyIsEmp(Dry) Then Exit Property
Dim Dr, O()
For Each Dr In Dry
   Push O, AyReOrd(Dr, PartialIxAy)
Next
DryReOrd = O
End Property

Property Get DryRmvColByIxAy(Dry, IxAy%()) As Variant()
If AyIsEmp(Dry) Then Exit Property
Dim O(), Dr
For Each Dr In Dry
   Push O, AyWhExclIxAy(Dr, IxAy)
Next
DryRmvColByIxAy = O
End Property

Property Get DryRowCnt&(Dry, ColIx&, EqVal)
If AyIsEmp(Dry) Then Exit Property
Dim J&, O&, Dr
For Each Dr In Dry
   If Dr(ColIx) = EqVal Then O = O + 1
Next
DryRowCnt = O
End Property

Property Get DrySel(A(), CIxAy&(), Optional CrtEmpCol_IfReqCol_NotFound As Boolean) As Variant()
Dim O(), Dr
If Sz(A) = 0 Then Exit Property
For Each Dr In A
   Push O, AyWhIxAy(Dr, CIxAy, CrtEmpCol_IfReqCol_NotFound)
Next
DrySel = O
End Property

Property Get DrySelDis(A(), ColIx%) As Variant()
If Sz(A) = 0 Then Exit Property
Dim Dr, O()
For Each Dr In A
   PushNoDup O, Dr(ColIx)
Next
DrySelDis = O
End Property

Property Get DrySelDisIntCol(Dry(), ColIx%) As Integer()
DrySelDisIntCol = AyIntAy(DrySelDis(Dry, ColIx))
End Property

Property Get DrySq(Dry, Optional NColOpt% = 0) As Variant()
If AyIsEmp(Dry) Then Exit Property
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
End Property

Property Get DrySrt(Dry, ColIx%, Optional IsDes As Boolean) As Variant()
Dim Col: Col = DryCol(Dry, ColIx)
Dim Ix&(): Ix = AySrtInToIxAy(Col, IsDes)
Dim J%, O()
For J = 0 To UB(Ix)
   Push O, Dry(Ix(J))
Next
DrySrt = O
End Property

Property Get DryStrCol(A, Optional ColIx% = 0) As String()
DryStrCol = DryCol_Into(A, ColIx, EmpSy)
End Property

Property Get DryWdtAy(A, Optional MaxColWdt& = 100) As Integer()
Const CSub$ = "DryWdtAy"
If Sz(A) = 0 Then Exit Property
Dim O%()
   Dim Dr, UDr%, U%, V, L%, J%
   U = -1
   For Each Dr In A
       If Not IsSy(Dr) Then Er CSub, "This routine should call DryCvFmtEachCell first so that each cell is ValCellStr as a string.|Now some Dr in given-Dry is not a StrAy, but[" & TypeName(Dr) & "]"
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
End Property

Property Get DryWh(A, ColIx%, EqVal) As Variant()
Dim O()
Dim J&
For J = 0 To UB(A)
   If A(J)(ColIx) = EqVal Then Push O, A(J)
Next
DryWh = O
End Property

Property Get DryWs(Dry, Optional WsNm$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = NewWs(WsNm, Vis:=True)
DryRg Dry, WsA1(O)
Set DryWs = O
End Property


Property Get DtDrpCol(A As Dt, Fny0) As Dt
Dim B As Drs: Set B = DtDrs(A)
Dim C As Drs: Set C = DrsDrpCol(B, Fny0)
Set DtDrpCol = Dt(A.DtNm, C.Fny, C.Dry)
End Property

Property Get DtDrs(A As Dt) As Drs
Set DtDrs = Drs(A.Fny, A.Dry)
End Property

Property Get DtIsEmp(A As Dt) As Boolean
DtIsEmp = AyIsEmp(A.Dry)
End Property

Property Get DtLy(A As Dt, Optional MaxColWdt& = 100, Optional BrkColNm$, Optional ShwZer As Boolean) As String()
Dim O$()
   Push O, "*Tbl " & A.DtNm
   PushAy O, DrsLy(DtDrs(A), MaxColWdt, BrkColNm, ShwZer)
DtLy = O
End Property

Property Get ItrCntByBoolPrp&(A, BoolPrpNm$)
If A.Count = 0 Then Exit Property
Dim O, Cnt&
For Each O In A
    If CallByName(O, BoolPrpNm, VbGet) Then
        Cnt = Cnt + 1
    End If
Next
ItrCntByBoolPrp = Cnt
End Property

Property Get ItrDrs(Itr, PrpNy0) As Drs
Dim Ny$()
    Ny = DftNy(PrpNy0)
Dim Dry()
    Dim Obj
    If Itr.Count > 0 Then
        For Each Obj In Itr
            Push Dry, ObjPrpDr(Obj, Ny)
        Next
    End If
Set ItrDrs = Drs(Ny, Dry)
End Property

Property Get ItrItmByPrp(A, PrpNm$, PrpV)
Dim O, V
If A.Count > 0 Then
    For Each O In A
        V = CallByName(O, PrpNm, VbGet)
        If V = PrpV Then
            Asg O, ItrItmByPrp
            Exit Property
        End If
    Next
End If
End Property

Property Get ItrNy(A, Optional Lik$ = "*") As String()
Dim O$(), Obj, N$
If A.Count > 0 Then
    For Each Obj In A
        N = Obj.Name
        If N Like Lik Then Push O, N
    Next
End If
ItrNy = O
End Property

Property Get ObjPrpDr(Obj, PrpNy0) As Variant()
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
End Property

Property Get S1S2Ay_Drs(A() As S1S2) As Drs
Set S1S2Ay_Drs = Drs("S1 S2", S1S2Ay_Dry(A))
End Property

Property Get S1S2Ay_Dry(A() As S1S2) As Variant()
Dim O()
Dim J%
For J = 0 To UB(A)
   With A(J)
       Push O, Array(.S1, .S2)
   End With
Next
S1S2Ay_Dry = O
End Property

Property Get SimTyStr_SimTy(SimTyStr$) As eSimTy
Dim O As eSimTy
Select Case UCase(SimTyStr)
Case "TXT": O = eTxt
Case "NBR": O = eNbr
Case "LGC": O = eLgc
Case "DTE": O = eDte
Case Else: O = eOth
End Select
SimTyStr_SimTy = O
End Property

Property Get SimTy_QuoteTp$(A As eSimTy)
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
End Property

Property Get SqNCol%(A)
On Error Resume Next
SqNCol = UBound(A, 2)
End Property

Property Get SqNRow%(A)
On Error Resume Next
SqNRow = UBound(A, 1)
End Property

Property Get SslDr(Ssl, TyAy() As eSimTy) As Variant()
Stop '
End Property

Property Get TitAy_Sq(TitAy$())
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
End Property

Property Get UIxAy(U&) As Long()
Dim O&(), J&
ReDim O(U)
For J = 0 To U
    O(J) = J
Next
UIxAy = O
End Property

Property Get VblLy_Dry(A$()) As Variant()
If Sz(A) = 0 Then Exit Property
Dim O()
   Dim I
   For Each I In A
       Push O, AyTrim(SplitVBar(CStr(I)))
   Next
VblLy_Dry = O
End Property

Sub DrsLoFmt(A As Drs, At As Range, LoFmtrLy$(), Optional LoNm$)
Dim Lo As ListObject
Stop '
'Set Lo = DrsLo(A, At, LoNm)
'LoFmt Lo, LoFmtrLy
End Sub

Sub DryBrw(Dry, Optional MaxColWdt& = 100, Optional BrkColIx% = -1)
AyBrw DryLy(Dry, MaxColWdt, BrkColIx)
End Sub

Sub DryDmp(Dry)
AyDmp DryLy(Dry)
End Sub

Function DryRg(A, At As Range) As Range
Set DryRg = SqRg(DrySq(A), At)
End Function

Sub DtBrw(A As Dt, Optional Fnn)
AyBrw DtLy(A), IIf(IsEmp(Fnn), A.DtNm, Fnn)
End Sub

Sub DtDmp(A As Dt)
AyDmp DtLy(A)
End Sub

Sub Fiy(Fny$(), FldLvs$, ParamArray OAp())
'Fiy=Field Index Array
Dim A$(): A = SplitSpc(FldLvs)
Dim I&(): I = AyIxAy(Fny, A)
Dim J%
For J = 0 To UB(I)
    OAp(J) = I(J)
Next
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

Private Property Get Dry_MgeIx&(Dry(), Dr, MgeIx%)
Dim O&, D, J%
For O = 0 To UB(Dry)
   D = Dry(O)
   For J = 0 To UB(Dr)
       If J <> MgeIx Then
           If Dr(J) <> D(J) Then GoTo Nxt
       End If
   Next
   Dry_MgeIx = O
   Exit Property
Nxt:
Next
Dry_MgeIx = -1
End Property

Private Sub ZZ_DrsSel()
'DrsBrw DrsSel(Vmd.MthDrs, "MthNm Mdy Ty MdNm")
'DrsBrw Vmd.MthDrs
End Sub

Private Sub ZZ_DsWb()
Dim Wb As Workbook
Stop '
'Set Wb = DsWb(DbDs(CurDb, "Permit PermitD"))
WbVis Wb
Stop
Wb.Close False
End Sub

Private Sub ZZ_ItrDrs()
Stop '
'DrsBrw ItrDrs(Dbt(SampleDb_DutyPrepare, "Permit").Flds, "Name Type Required")
'DrsBrw ItrDrs(Application.VBE.VBProjects, "Name Type")
End Sub

Private Sub ZZ_TitAy_Sq()
Dim A$()
Push A, "ksdf | skdfj  |skldf jf"
Push A, "skldf|sdkfl|lskdf|slkdfj"
Push A, "askdfj|sldkf"
Push A, "fskldf"
SqBrw TitAy_Sq(A)
End Sub

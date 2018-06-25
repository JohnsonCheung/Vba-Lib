Attribute VB_Name = "M_Drs"
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
Dim O As New Drs
Set ItrDrs = O.Init(Ny, Dry)
End Function

Private Sub ZZ_ItrDrs()
Stop
'DrsBrw ItrDrs(Dbt(SampleDb_DutyPrepare, "Permit").Flds, "Name Type Required")
'DrsBrw ItrDrs(Application.VBE.VBProjects, "Name Type")
End Sub

Property Get ToCellStr$(V, ShwZer As Boolean)
'CellStr is a string can be displayed in a cell
If QVb.M_Is.IsEmp(V) Then Exit Property
If IsStr(V) Then
    CellStr = V
    Exit Property
End If
If IsBool(V) Then
    CellStr = IIf(V, "TRUE", "FALSE")
    Exit Property
End If

If IsObject(V) Then
    CellStr = "[" & TypeName(V) & "]"
    Exit Property
End If
If ShwZer Then
    If IsNumeric(V) Then
        If V = 0 Then CellStr = "0"
        Exit Property
    End If
End If
If IsArray(V) Then
    If AyIsEmp(V) Then Exit Property
    CellStr = "Ay" & UB(V) & ":" & V(0)
    Exit Property
End If
If InStr(V, vbCrLf) > 0 Then
    CellStr = Brk(V, vbCrLf).S1 & "|.."
    Exit Property
End If
CellStr = V
End Property

Sub IxAy_Asg(IxAy%(), ParamArray OAp())
Dim J%
For J = 0 To UB(IxAy)
    If IsObject(OAp(J)) Then
        Stop '
        'Set OAp(J) = A(IxAy(J))
    Else
        Stop '
'        OAp(J) = A(IxAy(J))
    End If
Next
End Sub

Property Get SampleDrs() As Drs
Dim O As New Drs
Set SampleDrs = O.Init(SampleDrsFny, SampleDry)
End Property
Property Get URow&()
U = NRow - 1
End Property
Property Get IsEmp() As Boolean
IsEmp = N = 0
End Property
Property Get Lin$(Wdt%())
If IsEmp Then Exit Property
Dim UDr%
   UDr = U
Dim O$()
   Dim U1%: U1 = UB(Wdt)
   ReDim O(U1)
   Dim W, V
   Dim J%, V1$
   J = 0
   For Each W In Wdt
    Stop '
'       If UDr >= J Then V = A(J) Else V = ""
       V1 = AlignL(V, W)
       O(J) = V1
       J = J + 1
   Next
Lin = Quote(Join(O, " | "), "| * |")
End Property




Property Get Ly(Optional MaxColWdt& = 100, Optional BrkColNm$, Optional ShwZer As Boolean) As String()
'If BrkColNm changed, insert a break line
If AyIsEmp(A.Fny) Then Exit Property
Dim Drs As Drs
    Set Drs = AddRowIxCol
Dim BrkColIx%
    BrkColIx = AyIx(A.Fny, BrkColNm)
    If BrkColIx >= 0 Then BrkColIx = BrkColIx + 1 ' Need to increase by 1 due the Ix column is added
Dim Drx: Set Drx = Drs.Drx
Push Dry, Drs.Fny
Dim Ay$(): Ay = Dryx.Ly(MaxColWdt, BrkColIx:=BrkColIx, ShwZer:=ShwZer)  '<== Will insert break line if BrkColIx>=0
Dim Lin$: Lin = Pop(Ay)
Dim Hdr$: Hdr = Pop(Ay)
Dim O$()
    PushAy O, Array(Lin, Hdr)
    PushAy O, Ay
    Push O, Lin
Ly = O
End Property

Property Get InitByDic(A As Dictionary, Optional InclDicValTy As Boolean) As Drs
B_Fny = SplitSpc("Key Val"): If InclDicValTy Then Push B_Fny, "ValTy"
B_Dry = DrxByDic(A, InclDicValTy)
Set InitByDic = Me
End Property

Property Get AddConstCol(ColNm$, ConstVal) As Drs
Dim Fny$()
    Fny = B_Fny
    Push Fny, ColNm
Dim Dry()
    Dry = Drx.AddConstCol(ConstVal)
Dim O As New Drs
Set AddConstCol = O.Init(Fny, Dry)
End Property

Property Get AddRowIxCol() As Drs
Dim Fny$()
    Fny = AyIns(B_Fny, "Ix")
Dim Dry()
    If Not AyIsEmp(B_Dry) Then
        Dim J&, Dr
        For Each Dr In B_Dry
            Dr = AyIns(Dr, J): J = J + 1
            Push Dry, Dr
        Next
    End If
Dim O As New Drs
Set AddRowIxCol = O.Init(Fny, Dry)
End Property

Sub Brw(Optional MaxColWdt& = 100, Optional BrkColNm$, Optional Fnn$)
AyBrw Ly(MaxColWdt, BrkColNm$), Fnn
End Sub

Property Get Col(ColNm$) As Variant()
Col = Drx.Col(AyIx(B_Fny, ColNm))
End Property

Sub Dmp(Optional MaxColWdt& = 100, Optional BrkColNm$)
AyDmp Ly(MaxColWdt, BrkColNm$)
End Sub

Property Get DrpCol(ColNy0) As Drs
Dim ColNy$(): ColNy = DftNy(ColNy0)
Ass AyHasSubAy(B_Fny, ColNy)
Dim IxAy%()
    IxAy = FnyIxAy(B_Fny, ColNy)
Dim Fny$(), Dry()
    Fny = AyWhExclIxAy(B_Fny, IxAy)
    Dry = Drx.RmvColByIxAy(IxAy)
Set DrpCop = Drs(Fny, Dry)
End Property

Property Get ExpLinesCol(LinesColNm$) As Drs
Dim Ix%
    Ix = AyIx(B_Fny, LinesColNm)
Dim Dry()
    Dim Dr
    For Each Dr In B_Dry
        PushAy Dry, DrExpLinesCol(Dr, Ix)
    Next
DrsExpLinesCol = Drs(B_Fny, Dry)
End Property

Property Get FldLvs$(A As Drs)
DrsFldLvs = JnSpc(A.Fny)
End Property

Property Get InitByLo(A As ListObject) As Drs
With LoDrs
    .Dry = LoDry(A)
    .Fny = LoFny(A)
End With
End Property
Function Lo(At As Range, Optional LoNm$, Optional StopAutoFit As Boolean) As ListObject
AyRgH A.Fny, At
Dim Rg As Range: Set Rg = Drx.Rg(RgRC(At, 2, 1))
Dim R1 As Range: Set R1 = RgRR(Rg, 0, RgNRow(Rg))
Set DrsLo = RgLo(R1, LoNm)
If Not StopAutoFit Then
    '\At Fny->AutoFit
    Dim R2 As Range: Set R2 = RgCC(At, 1, Sz(A.Fny)).EntireColumn
    R2.AutoFit
End If
End Function

Function Ws(Optional WsNm$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = NewWs(WsNm, Vis:=True)
Lo WsA1(O)
Set Ws = O
End Function

Function LoWithFmt(At As Range, LoFmtrLy$(), Optional LoNm$) As ListObject
Dim Lo As ListObject
Set Lo = Me.Lo(At, LoNm)
'LoFmt Lo, LoFmtrLy
End Function

Property Get ReOrd(Partial_Fny0) As Drs
Dim ReOrdFny$(): ReOrdFny = DftNy(Partial_Fny0)
Dim IxAy&(): IxAy = AyIxAy(A.Fny, ReOrdFny)
Dim OFny$(): OFny = AyReOrd(A.Fny, IxAy)
Dim ODry(): ODry = Drx.ReOrd(IxAy)
Dim O As New Drs
Set ReOrd = O.Init(OFny, ODry)
End Property

Property Get RowCnt&(ColNm$, EqVal)
RowCnt = Drx.RowCnt(AyIx(A.Fny, ColNm), EqVal)
End Property

Property Get DrsSel(Fny0, Optional CrtEmpColIfReqFldNotFound As Boolean) As Drs
Dim Fny$(): Fny = DftNy(Fny0)
Dim IxAy&()
    If CrtEmpColIfReqFldNotFound Then
        IxAy = AyIxAy(A.Fny, Fny, ChkNotFound:=False)
    Else
        IxAy = AyIxAy(A.Fny, Fny, ChkNotFound:=True)
    End If
Dim O As New Drs
Set Sel = O.Init(Fny, Drx.Sel(IxAy, CrtEmpColIfReqFldNotFound))
End Property

Property Get Srt(ColNm$, Optional IsDes As Boolean) As Drs
DrsSrt = Drs(B_Fny, Drx.Srt(AyIx(B_Fny, ColNm), IsDes))
End Property

Property Get StrCol(ColNm$) As String()
StrCol = AySy(Col(ColNm))
End Property

Property Get Wh(Fld, V) As Drs
Const CSub$ = "Drs.Wh"
Dim Ix%:
    Ix = AyIx(A.Fny, Fld)
    If Ix = -1 Then Er CSub, "{Fld} is not in {Fny}", Fld, A.Fny
Set Wh = Drs(B_Fny, Drx.Wh(Ix, V))
End Property

Property Get WhNotRowIxAy(RowIxAy&()) As Drs
Dim O()
    Dim J&
    For J = 0 To URow
        If Not AyHas(RowIxAy, J) Then
            Push O, B_Dry(J)
        End If
    Next
Set WhNotRowIxAy = Drs(B_Fny, O)
End Property

Property Get WhRow(RowIxAy&()) As Drs
Dim O()
    If Not AyIsEmp(RowIxAy) Then
        Dim I
        For Each I In RowIxAy
            Push O, B_Dry(I)
        Next
    End If
Set WhRow = Drs(B_Fny, O)
End Property

Private Sub ZZ_Sel()
Stop
'Sel(CurPjx.Mths.Drs, "MthNm Mdy Ty MdNm").Brw
End Sub

Property Get InitByDicAy(DicAy, Optional Fny0) As Drs
Const CSub$ = "InitByDicAy"
Dim UDic%
   UDic = UB(DicAy)

Erase B_Fny
   If AyIsEmp(Fny0) Then
       Dim J%
       Push B_Fny, "Key"
       For J = 0 To UDic
           Push B_Fny, "V" & J
       Next
   Else
       B_Fny = Fny0
   End If
If UB(B_Fny) <> UDic + 1 Then Er CSub, "Given {Fny0} has {Sz} <> {DicAy-Sz}", FnyOpt, Sz(FnyOpt), Sz(DicAy)
Dim Ky()
   Ky = DicAy_Ky(DicAy)
Dim URow&
   URow = UB(Ky)
Dim O()
   ReDim O(URow)
   Dim K
   J = 0
   For Each K In Ky
       O(J) = DicAy_Dr(DicAy, K)
       J = J + 1
   Next
DicJn.Dry = O
B_Fny = Fny
Set InitByDicAy = Me
End Property

Property Get InitByLines(DrsLines$) As Drs
Set InitByLines = InitByLy(SplitCrLf(DrsLines))
End Property

Property Get InitByLy(DrsLy$()) As Drs
Dim Fny$(): Fny = SslSy(DrsLy(0))
Dim J&, Dry()
If IsSimTyLvs(DrsLy(2)) Then
    Dim TyAy() As eSimTy
    For J = 3 To UB(DrsLy)
        Push Dry, DrBySsl(DrsLy(J), TyAy)
    Next
Else
    For J = 2 To UB(DrsLy)
        Push Dry, SslSy(DrsLy(J))
    Next
End If
Set InitByLy = Init(Fny, Dry)
End Property

Property Get InitByVbl(DrsVbl$) As Drs
'SpecStr:Vbl:VbarLine
'SpecStr:DrsVbl:Data-record-set-vbar-line
Set InitByVbl = InitByLy(SplitVBar(DrsVbl))
End Property



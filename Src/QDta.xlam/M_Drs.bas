Attribute VB_Name = "M_Drs"
Option Explicit
Property Get ItrDrs(Itr, PrpNy0) As Drs
Dim Ny$()
    Ny = DftNy(PrpNy0)
Dim Dry()
    Dim Obj
    If Itr.Count > 0 Then
        For Each Obj In Itr
            Stop '
'            Push Dry, ObjPrpDr(Obj, Ny)
        Next
    End If
Dim O As New Drs
Set ItrDrs = O.Init(Ny, Dry)
End Property

Private Sub ZZ_ItrDrs()
Stop
'DrsBrw ItrDrs(Dbt(SampleDb_DutyPrepare, "Permit").Flds, "Name Type Required")
'DrsBrw ItrDrs(Application.VBE.VBProjects, "Name Type")
End Sub

Property Get ToCellStr$(V, ShwZer As Boolean)
'CellStr is a string can be displayed in a cell
If QVb.M_Is.IsEmp(V) Then Exit Property
If IsStr(V) Then
    ToCellStr = V
    Exit Property
End If
If IsBool(V) Then
    ToCellStr = IIf(V, "TRUE", "FALSE")
    Exit Property
End If

If IsObject(V) Then
    ToCellStr = "[" & TypeName(V) & "]"
    Exit Property
End If
If ShwZer Then
    If IsNumeric(V) Then
        If V = 0 Then ToCellStr = "0"
        Exit Property
    End If
End If
If IsArray(V) Then
    If AyIsEmp(V) Then Exit Property
    ToCellStr = "Ay" & UB(V) & ":" & V(0)
    Exit Property
End If
If InStr(V, vbCrLf) > 0 Then
    ToCellStr = Brk(V, vbCrLf).S1 & "|.."
    Exit Property
End If
ToCellStr = V
End Property

Property Get DrLin$(A, Wdt%())
If Sz(A) = 0 Then Exit Property
Dim UDr%
   UDr = UB(A)
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
DrLin = Quote(Join(O, " | "), "| * |")
End Property

Property Get DrsLy(A As Drs, Optional MaxColWdt& = 100, Optional BrkColNm$, Optional ShwZer As Boolean) As String()
'If BrkColNm changed, insert a break line
If Sz(A.Fny) = 0 Then Exit Property
Dim Drs As Drs
    Set Drs = DrsAddRowIxCol(A)
Dim BrkColIx%
    BrkColIx = AyIx(A.Fny, BrkColNm)
    If BrkColIx >= 0 Then BrkColIx = BrkColIx + 1 ' Need to increase by 1 due the Ix column is added
Dim Dry(): Dry = A.Dry
Push Dry, A.Fny
Dim Ay$(): Ay = DryLy(Dry, MaxColWdt, BrkColIx:=BrkColIx, ShwZer:=ShwZer) '<== Will insert break line if BrkColIx>=0
Dim Lin$: Lin = Pop(Ay)
Dim Hdr$: Hdr = Pop(Ay)
Dim O$()
    PushAy O, Array(Lin, Hdr)
    PushAy O, Ay
    Push O, Lin
DrsLy = O
End Property

Property Get DicDrs(A As Dictionary, Optional InclDicValTy As Boolean) As Drs
Dim Fny$()
Fny = SplitSpc("Key Val"): If InclDicValTy Then Push Fny, "ValTy"
Set DicDrs = Drs(Fny, DicDry(A, InclDicValTy))
End Property

Property Get DrsAddConstCol(A As Drs, ColNm$, ConstVal) As Drs
Dim Fny$()
    Fny = A.Fny
    Push Fny, ColNm
Set DrsAddConstCol = Drs(Fny, DryAddConstCol(A.Dry, ConstVal))
End Property

Property Get DrsAddRowIxCol(A As Drs) As Drs
Dim Fny$()
    Fny = AyIns(A.Fny, "Ix")
Dim Dry()
    If Not AyIsEmp(A.Dry) Then
        Dim J&, Dr
        For Each Dr In A.Dry
            Dr = AyIns(Dr, J): J = J + 1
            Push Dry, Dr
        Next
    End If
Set DrsAddRowIxCol = Drs(Fny, Dry)
End Property

Sub DrsBrw(A As Drs, Optional MaxColWdt& = 100, Optional BrkColNm$, Optional Fnn$)
AyBrw DrsLy(A, MaxColWdt, BrkColNm$), Fnn
End Sub

Property Get DrsCol(A As Drs, ColNm$) As Variant()
DrsCol = DryCol(A.Dry, AyIx(A.Fny, ColNm))
End Property

Sub DrsDmp(A As Drs, Optional MaxColWdt& = 100, Optional BrkColNm$)
AyDmp DrsLy(A, MaxColWdt, BrkColNm$)
End Sub

Property Get DrsDrpCol(A As Drs, ColNy0) As Drs
Dim ColNy$(): ColNy = DftNy(ColNy0)
Ass AyHasSubAy(A.Fny, ColNy)
Dim IxAy%()
    IxAy = FnyIxAy(A.Fny, ColNy)
Dim Fny$(), Dry()
    Fny = AyWhExclIxAy(A.Fny, IxAy)
    Dry = DryRmvColByIxAy(A.Dry, IxAy)
Set DrsDrpCol = Drs(Fny, Dry)
End Property

Property Get DrsExpLinesCol(A As Drs, LinesColNm$) As Drs
Dim Dry(): Dry = A.Dry
If Sz(Dry) = 0 Then
    Set DrsExpLinesCol = Drs(A.Fny, Dry)
    Exit Property
End If
Dim Ix%
    Ix = AyIx(A.Fny, LinesColNm)
Dim O()
    Dim Dr
    For Each Dr In A.Dry
        Stop 'sotp
        'PushAy Dry, DrExpLinesCol(Dr, Ix)
    Next
DrsExpLinesCol = Drs(A.Fny, O)
End Property

Property Get DrsFldSsl$(A As Drs)
DrsFldSsl = JnSpc(A.Fny)
End Property

Property Get LoDrs(A As ListObject) As Drs
Set LoDrs = Drs(LoFny(A), LoDry(A))
End Property

Function DrsLo(A As Drs, At As Range, Optional LoNm$, Optional StopAutoFit As Boolean) As ListObject
AyRgH A.Fny, At
Dim Rg As Range: Set Rg = DryRg(A.Dry, RgRC(At, 2, 1))
Dim R1 As Range: Set R1 = RgRR(Rg, 0, RgNRow(Rg))
Set DrsLo = RgLo(R1, LoNm)
If Not StopAutoFit Then
    '\At Fny->AutoFit
    Dim R2 As Range: Set R2 = RgCC(At, 1, Sz(A.Fny)).EntireColumn
    R2.AutoFit
End If
End Function

Property Get DrsWs(A As Drs, Optional WsNm$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = NewWs(WsNm, Vis:=True)
DrsLo A, WsA1(O)
Set DrsWs = O
End Property

Function DrsLoWithFmt(A As Drs, At As Range, LoFmtrLy$(), Optional LoNm$) As ListObject
Dim Lo As ListObject
Set Lo = DrsLo(A, At, LoNm)
Stop '
'LoFmt Lo, LoFmtrLy
End Function

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
Dim O As New Drs
Set DrsSel = Drs(Fny, DrySel(A.Dry, IxAy, CrtEmpColIfReqFldNotFound))
End Property


Property Get DrsStrCol(A As Drs, ColNm$) As String()
DrsStrCol = AySy(DrsCol(A, ColNm))
End Property

Property Get DrsWh(A As Drs, Fld, V) As Drs
Set DrsWh = Drs(A.Fny, DryWh(A.Dry, AyIx(A.Fny, Fld), V))
End Property

Property Get DrsWhNotRowIxAy(A As Drs, RowIxAy&()) As Drs
Dim O(), Dry()
    Dry = A.Dry
    Dim J&
    For J = 0 To UB(Dry)
        If Not AyHas(RowIxAy, J) Then
            Push O, Dry(J)
        End If
    Next
Set DrsWhNotRowIxAy = Drs(A.Fny, O)
End Property

Property Get DrsWhRow(A As Drs, RowIxAy&()) As Drs
Dim O(), Dry()
    Dry = A.Dry
    If Not AyIsEmp(RowIxAy) Then
        Dim I
        For Each I In RowIxAy
            Push O, Dry(I)
        Next
    End If
Set DrsWhRow = Drs(A.Fny, O)
End Property

Private Sub ZZ_DrsSel()
DrsBrw DrsSel(SampleDrs, "A B D")
End Sub

Property Get DicAy_Drs(A() As Dictionary, Optional Fny0) As Drs
Const CSub$ = "DicAy_Drs"
Dim UDic%
   UDic = UB(A)
Dim Fny$()
    Fny = DftNy(Fny0)
    If AyIsEmp(Fny) Then
        Dim J%
        Push Fny, "Key"
        For J = 0 To UDic
            Push Fny, "V" & J
        Next
    Else
        Fny = Fny0
    End If
If UB(Fny) <> UDic + 1 Then Er CSub, "Given {Fny0} has {Sz} <> {DicAy-Sz}", Fny, Sz(Fny), Sz(A)
Dim Ky$()
   Ky = DicAy_Ky(A)
Dim O()
    Dim I
   ReDim O(UDic)
   J = 0
   For Each I In A
       O(J) = ZDicDr(CvDic(I), Ky)
       J = J + 1
   Next
Set DicAy_Drs = Drs(Fny, O)
End Property

Property Get DicAy_Ky(A() As Dictionary) As String()
Dim O$(), I
For Each I In A
    PushNoDupAy O, CvDic(I).Keys
Next
DicAy_Ky = O
End Property

Private Property Get ZDicDr(A As Dictionary, Ky$()) As Variant()
Dim O(), I, J&
ReDim O(UB(Ky))
For Each I In Ky
    If A.Exists(I) Then
        O(J) = A(I)
    End If
    J = J + 1
Next
ZDicDr = O
End Property
Property Get DrsLines_Drs(A) As Drs
Set DrsLines_Drs = DrsLy_Drs(SplitLines(A))
End Property

Property Get DrsLy_Drs(A$()) As Drs
Dim Fny$(): Fny = SslSy(A(0))
Dim J&, Dry()
If IsSimTySsl(A(2)) Then
    Dim TyAy() As eSimTy
    For J = 3 To UB(A)
        Push Dry, SslDr(A(J), TyAy)
    Next
Else
    For J = 2 To UB(A)
        Push Dry, SslSy(A(J))
    Next
End If
Set DrsLy_Drs = Drs(Fny, Dry)
End Property

Property Get DrsVbl_Drs(A) As Drs
'SpecStr:Vbl:VbarLine
'SpecStr:DrsVbl:Data-record-set-vbar-line
Set DrsVbl_Drs = DrsLy_Drs(SplitVBar(A))
End Property



Attribute VB_Name = "M_Drs"
Option Explicit

Function DrsAddConstCol(A As Drs, ColNm$, ConstVal) As Drs
Dim Fny$()
    Fny = A.Fny
    Push Fny, ColNm
Set DrsAddConstCol = Drs(Fny, DryAddConstCol(A.Dry, ConstVal))
End Function

Function DrsAddRowIxCol(A As Drs) As Drs
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
End Function

Function DrsCol(A As Drs, ColNm$) As Variant()
DrsCol = DryCol(A.Dry, AyIx(A.Fny, ColNm))
End Function

Function DrsDrpCol(A As Drs, ColNy0) As Drs
Dim ColNy$(): ColNy = DftNy(ColNy0)
Ass AyHasSubAy(A.Fny, ColNy)
Dim IxAy%()
    IxAy = FnyIxAy(A.Fny, ColNy)
Dim Fny$(), Dry()
    Fny = AyWhExclIxAy(A.Fny, IxAy)
    Dry = DryRmvColByIxAy(A.Dry, IxAy)
Set DrsDrpCol = Drs(Fny, Dry)
End Function

Function DrsDt(A As Drs, DtNm$) As Dt
Set DrsDt = Dt(DtNm, A.Fny, A.Dry)
End Function

Function DrsExpLinesCol(A As Drs, LinesColNm$) As Drs
Dim Dry(): Dry = A.Dry
If Sz(Dry) = 0 Then
    Set DrsExpLinesCol = Drs(A.Fny, Dry)
    Exit Function
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
End Function

Function DrsFldSsl$(A As Drs)
DrsFldSsl = JnSpc(A.Fny)
End Function

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

Function DrsLoWithFmt(A As Drs, At As Range, LoFmtrLy$(), Optional LoNm$) As ListObject
Dim Lo As ListObject
Set Lo = DrsLo(A, At, LoNm)
Stop '
'LoFmt Lo, LoFmtrLy
End Function

Function DrsLy(A As Drs, Optional MaxColWdt& = 100, Optional BrkColNm$, Optional ShwZer As Boolean) As String()
'If BrkColNm changed, insert a break line
If Sz(A.Fny) = 0 Then Exit Function
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
Dim O As New Drs
Set DrsSel = Drs(Fny, DrySel(A.Dry, IxAy, CrtEmpColIfReqFldNotFound))
End Function

Function DrsSrt(A As Drs, ColNm$, Optional IsDes As Boolean) As Drs
Set DrsSrt = Drs(A.Fny, DrySrt(A.Dry, AyIx(A.Fny, ColNm), IsDes))
End Function

Function DrsStrCol(A As Drs, ColNm$) As String()
DrsStrCol = AySy(DrsCol(A, ColNm))
End Function

Function DrsWh(A As Drs, Fld, V) As Drs
Set DrsWh = Drs(A.Fny, DryWh(A.Dry, AyIx(A.Fny, Fld), V))
End Function

Function DrsWhNotRowIxAy(A As Drs, RowIxAy&()) As Drs
Dim O(), Dry()
    Dry = A.Dry
    Dim J&
    For J = 0 To UB(Dry)
        If Not AyHas(RowIxAy, J) Then
            Push O, Dry(J)
        End If
    Next
Set DrsWhNotRowIxAy = Drs(A.Fny, O)
End Function

Function DrsWhRow(A As Drs, RowIxAy&()) As Drs
Dim O(), Dry()
    Dry = A.Dry
    If Not AyIsEmp(RowIxAy) Then
        Dim I
        For Each I In RowIxAy
            Push O, Dry(I)
        Next
    End If
Set DrsWhRow = Drs(A.Fny, O)
End Function

Function DrsWhRowIxAy(A As Drs, RowIxAy&()) As Drs
Dim O()
    Dim J&, I, Dry()
    Dry = A.Dry
    For Each I In RowIxAy
        Push O, Dry(I)
    Next
Set DrsWhRowIxAy = Drs(A.Fny, O)
End Function

Function DrsWs(A As Drs, Optional WsNm$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = NewWs(WsNm, Vis:=True)
DrsLo A, WsA1(O)
Set DrsWs = O
End Function

Sub DrsBrw(A As Drs, Optional MaxColWdt& = 100, Optional BrkColNm$, Optional Fnn$)
AyBrw DrsLy(A, MaxColWdt, BrkColNm$), Fnn
End Sub

Sub DrsDmp(A As Drs, Optional MaxColWdt& = 100, Optional BrkColNm$)
AyDmp DrsLy(A, MaxColWdt, BrkColNm$)
End Sub

Sub DrsLoFmt(A As Drs, At As Range, LoFmtrLy$(), Optional LoNm$)
Dim Lo As ListObject
Stop '
'Set Lo = DrsLo(A, At, LoNm)
'LoFmt Lo, LoFmtrLy
End Sub

Private Sub ZZ_DrsSel()
DrsBrw DrsSel(SampleDrs, "A B D")
End Sub

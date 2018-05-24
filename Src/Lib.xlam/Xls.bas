Attribute VB_Name = "Xls"
Option Explicit
Public Const SampleFx_KE24 = "C:\Users\User\Desktop\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls"

Function AyRgH(Ay, At As Range) As Range
Set AyRgH = SqRg(AySqH(Ay), At)
End Function

Function AyRgV(Ay, At As Range) As Range
Set AyRgV = SqRg(AySqV(Ay), At)
End Function

Function CellAyH(A As Range) As Variant()
If IsEmpty(RgA1(A).Value) Then Exit Function
Dim R&
For R = 2 To MaxRow
    If IsEmpty(RgRC(A, R, 1).Value) Then
        CellAyH = SqCol(RgCRR(A, 1, 1, R - 1).Value, 1)
        Exit Function
    End If
Next
Stop
End Function

Function CellAyV(A As Range) As Variant()

End Function

Sub CellClrDown(Cell As Range)
CellVBar(Cell, AtLeastOneCell:=True).Clear
End Sub

Sub CellFillSeqDown(Cell As Range, N&, Optional IsFmOne As Boolean)
AyRgV NewIntSeq(N, IsFmOne), Cell
End Sub

Function CellIsInRg(A As Range, Rg As Range) As Boolean
Dim R&, C%, R1&, R2&, C1%, C2%
R = A.Row
R1 = Rg.Row
If R < R1 Then Exit Function
R2 = R1 + Rg.Rows.Count
If R > R2 Then Exit Function
C = A.Column
C1 = Rg.Column
If C < C1 Then Exit Function
C2 = C1 + Rg.Columns.Count
If C > C2 Then Exit Function
CellIsInRg = True
End Function

Function CellIsInRgAp(A As Range, ParamArray RgAp()) As Boolean
Dim Av(): Av = RgAp
CellIsInRgAp = CellIsInRgAv(A, Av)
End Function

Function CellIsInRgAv(A As Range, RgAv()) As Boolean
Dim V
For Each V In RgAv
    If CellIsInRg(A, CvRg(V)) Then CellIsInRgAv = True: Exit Function
Next
End Function

Sub CellLnkWs(A As Range, WsNy$())
Dim WsNm: WsNm = A.Value
If Not VarIsStr(WsNm) Then Exit Sub
If Not AyHas(WsNy, WsNm) Then Exit Sub
With A.Hyperlinks
    If .Count > 0 Then .Delete
    .Add A, "", FmtQQ("'?'!A1", WsNm)
End With
End Sub

Function CellVBar(Cell As Range, Optional AtLeastOneCell As Boolean) As Range
If IsEmpty(Cell.Value) Then
    If AtLeastOneCell Then
        Set CellVBar = RgA1(Cell)
    End If
    Exit Function
End If
Dim R1&: R1 = Cell.Row
Dim R2&
    If IsEmpty(RgRC(Cell, 2, 1)) Then
        R2 = Cell.Row
    Else
        R2 = Cell.End(xlDown).Row
    End If
Dim C%: C = Cell.Column
Set CellVBar = WsCRR(RgWs(Cell), C, R1, R2)
End Function

Function CurWb() As Workbook
Set CurWb = Excel.Application.ActiveWorkbook
End Function

Function CvLo(A) As ListObject
Set CvLo = A
End Function

Function CvLoCol(A) As ListColumn
Set CvLoCol = A
End Function

Function CvRg(A) As Range
Set CvRg = A
End Function

Function DftFx$(A$)
If A = "" Then
   Dim O$: O = TmpFx
   DftFx = O
Else
   DftFx = A
End If
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

Function DrsWs(A As Drs, Optional WsNm$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = NewWs(WsNm, Vis:=True)
DrsLo A, WsA1(O)
Set DrsWs = O
End Function

Function DryRg(Dry, At As Range) As Range
Set DryRg = SqRg(DrySq(Dry), At)
End Function

Function DsNDt%(A As Ds)
DsNDt = DtAySz(A.DtAy)
End Function

Function DsWb(A As Ds, Optional Vis As Boolean) As Workbook
Dim O As Workbook
Set O = NewWb
With WbFstWs(O)
   .Name = "Ds"
   .Range("A1").Value = A.DsNm
End With
If Not DsIsEmp(A) Then
   Dim J%
   For J = 0 To DsNDt(A) - 1
       WbAddDt O, A.DtAy(J)
   Next
End If
If Vis Then WbVis O
Set DsWb = O
End Function

Function DsWs(A As Ds, Optional Hid As Boolean) As Worksheet
Dim O As Worksheet: Set O = NewWs(Vis:=Not Hid)
WsA1(O).Value = "*Ds " & A.DsNm
Dim At As Range, J%
Set At = WsRC(O, 2, 1)
For J = 0 To DsNDt(A)
    Set At = DtAt_NxtAt(A.DtAy(J), At, J)
Next
DsWs = O
End Function

Function DtAt_NxtAt(A As Dt, At As Range, J%) As Range
At.Value = "(" & J & ") " & A.DtNm
Set At = RgRC(At, 2, 1)
Dim Ly$(): Ly = DrsLy(DtDrs(A))
AyRgV Ly, At
Set At = RgRC(At, 1 + Sz(Ly), 1)
Set DtAt_NxtAt = At
End Function

Function DtLo(A As Dt, At As Range) As ListObject
Dim R As Range
If At.Row = 1 Then
    Set R = RgRC(At, 2, 1)
Else
    Set R = At
End If
Set DtLo = DrsLo(DtDrs(A), R, A.DtNm)
RgRC(R, 0, 1).Value = A.DtNm
End Function

Function DtWs(A As Dt, Optional Vis As Boolean) As Worksheet
Dim O As Worksheet
Set O = NewWs(A.DtNm)
DrsLo DtDrs(A), WsA1(O)
Set DtWs = O
If Vis Then WsVis O
End Function

Function FxAcnStr$(A)
FxAcnStr = FmtQQ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=?;Extended Properties=""Excel 12.0;HDR=YES""", A)
End Function

Function FxCn(A) As ADODB.Connection
Dim O As New ADODB.Connection
O.Open FxAcnStr(A)
Set FxCn = O
End Function

Function FxHasWs(A$, WsNm) As Boolean
FxHasWs = AyHas(FxWsNy(A), WsNm)
End Function

Function FxOpn(A$) As Workbook
Set FxOpn = NewXls.Workbooks.Open(A)
End Function

Sub FxRmvWsIfExist(A$, WsNm)
If FxHasWs(A, WsNm) Then
   Dim B As Workbook: Set B = FxOpn(A)
   WbWs(B, WsNm).Delete
   WbSav B
   WbClsNoSav B
End If
End Sub

Sub FxSql_Arun(A$, Sql$)
FxCn(A).Execute Sql
End Sub

Function FxWsFny(A$, WsNm$) As String()
Dim Cat As ADOX.Catalog: Set Cat = FxCat(A)
Dim C As ADOX.Column
Dim O$()
Dim N$: N = WsNm & "$"
For Each C In Cat.Tables(N).Columns
   Push O, C.Name
Next
FxWsFny = O
End Function

Function FxaCrt(Optional Fxa0$) As Excel.Application
Dim Fxa$: Fxa = DftFfn(Fxa0, ".xlam")
Dim Wb As Workbook, X As Excel.Application
   Set X = New Excel.Application
   Set Wb = X.Workbooks.Add
X.VBE.VBProjects(1).Name = FfnFnn(Fxa)
Wb.SaveAs Fxa, XlFileFormat.xlOpenXMLAddIn
Wb.Close
X.AddIns.Add Fxa
Set FxaCrt = X
End Function

Sub HBar_MgeSamValCell(A As Range)
Ass RgIsHBar(A)
Dim R As Range
Set R = HBar_SamValRg(A)
Dim Sav As Boolean
    Sav = A.Application.DisplayAlerts
    A.Application.DisplayAlerts = False
While Not IsNothing(R)
    R.Merge '<===================================
    Set R = HBar_SamValRg(R)
Wend
A.Application.DisplayAlerts = Sav
End Sub

Function HBar_SamValRg(A As Range) As Range
Dim NCol%: NCol = RgNCol(A)
Dim C1%, V, C2%, Fnd As Boolean
For C1 = 1 To NCol - 1
    V = RgRC(A, 1, C1).Value
    For C2 = C1 + 1 To NCol
        If RgRC(A, 1, C2).Value = V Then
            Fnd = True
        Else
            If Fnd Then
                C2 = C2 - 1
                GoTo Fnd
            End If
            GoTo Nxt
        End If
    Next
Nxt:
Next
Fnd:
If Fnd Then Set HBar_SamValRg = RgRCC(A, 1, C1, C2)
End Function

Function IsNum(A) As Boolean
Dim J%
For J = 1 To Len(A)
    If Not IsDigit(Mid(A, J, 1)) Then Exit Function
Next
IsNum = True
End Function

Sub LoAdjColWdt(A As ListObject)
Dim C As Range: Set C = LoEntCol(A)
C.AutoFit
Dim EntC As Range, J%
For J = 1 To C.Columns.Count
   Set EntC = RgEntC(C, J)
   If EntC.ColumnWidth > 100 Then EntC.ColumnWidth = 100
Next
End Sub

Sub LoBrw(A As ListObject)
DrsBrw LoDrs(A)
End Sub

Function LoC(A As ListObject, C, Optional InclTot As Boolean, Optional InclHdr As Boolean) As Range
Dim mC%, R1&, R2&
R1 = LoR1(A, InclHdr)
R2 = LoR2(A, InclTot)
C = LoWsCno(A, C)
Set LoC = WsCRR(LoWs(A), mC, R1, R2)
End Function

Function LoCC(A As ListObject, C1, C2, Optional InclTot As Boolean, Optional InclHdr As Boolean) As Range
Dim R1&, R2&, mC1%, mC2%
R1 = LoR1(A, InclHdr)
R2 = LoR2(A, InclTot)
mC1 = LoWsCno(A, C1)
mC2 = LoWsCno(A, C2)
Set LoCC = WsRCRC(LoWs(A), R1, mC1, R2, mC2)
End Function

Sub LoColNm_SetAvg(A As ListObject, F)
LoColNm_SetSummary A, F, xlTotalsCalculationAverage
End Sub

Sub LoColNm_SetCnt(A As ListObject, F)
LoColNm_SetSummary A, F, xlTotalsCalculationCount
End Sub

Sub LoColNm_SetSummary(A As ListObject, F, Tot As XlTotalsCalculation)
A.ListColumns(F).TotalsCalculation = Tot
End Sub

Sub LoColNm_SetTot(A As ListObject, F)
LoColNm_SetSummary A, F, xlTotalsCalculationSum
End Sub

Sub LoCol_LnkWs(A As ListObject, ColNm$)
RgLnkWs LoCol_Rg(A, ColNm)
End Sub

Function LoCol_Rg(A As ListObject, ColNm$) As Range
Set LoCol_Rg = A.ListColumns(ColNm).Range
End Function

Function LoCrt(A As Worksheet, Optional LoNm$) As ListObject
Dim R As Range: Set R = WsDtaRg(A)
If IsNothing(R) Then Exit Function
Dim O As ListObject: Set O = A.ListObjects.Add(xlSrcRange, WsDtaRg(A), , xlYes)
If LoNm <> "" Then O.Name = LoNm
LoAdjColWdt O
Set LoCrt = O
End Function

Function LoDrs(A As ListObject) As Drs
With LoDrs
    .Dry = LoDry(A)
    .Fny = LoFny(A)
End With
End Function

Function LoDry(A As ListObject) As Variant()
LoDry = SqDry(LoSq(A))
End Function

Function LoEntCol(A As ListObject) As Range
Set LoEntCol = LoCC(A, 1, LoNCol(A)).EntireColumn
End Function

Function LoFny(A As ListObject) As String()
Dim O$(), I
For Each I In A.ListColumns
    Push O, CvLoCol(I).Name
Next
LoFny = O
End Function

Function LoHasNoDta(A As ListObject) As Boolean
LoHasNoDta = IsNothing(A.DataBodyRange)
End Function

Function LoHdrCell(A As ListObject, FldNm) As Range
Dim Rg As Range: Set Rg = A.ListColumns(FldNm).Range
Set LoHdrCell = RgRC(Rg, 1, 1)
End Function

Function LoNCol%(A As ListObject)
LoNCol = A.ListColumns.Count
End Function

Function LoR1&(A As ListObject, Optional InclHdr As Boolean)
If LoHasNoDta(A) Then
   LoR1 = A.ListColumns(1).Range.Row + 1
   Exit Function
End If
LoR1 = A.DataBodyRange.Row - IIf(InclHdr, 1, 0)
End Function

Function LoR2&(A As ListObject, Optional InclTot As Boolean)
If LoHasNoDta(A) Then
   LoR2 = LoR1(A)
   Exit Function
End If
LoR2 = A.DataBodyRange.Row + IIf(InclTot, 1, 0)
End Function

Function LoSq(A As ListObject)
LoSq = A.DataBodyRange.Value
End Function

Sub LoVis(A As ListObject)
A.Application.Visible = True
End Sub

Function LoWs(A As ListObject) As Worksheet
Set LoWs = A.Parent
End Function

Function LoWsCno%(A As ListObject, Ix_or_ColNm)
LoWsCno = A.ListColumns(Ix_or_ColNm).Range.Column
End Function

Function LyWs(Ly$(), Vis As Boolean) As Worksheet
Dim O As Worksheet: Set O = NewWs(Vis:=Vis)
AyRgV Ly, WsA1(O)
Set LyWs = O
End Function

Function MaxCol%()
Static C%
If C = 0 Then
    Dim Ws As Worksheet
    Set Ws = ActiveSheet
    C = Ws.Cells.Columns.Count
End If
MaxCol = C
End Function

Function MaxRow&()
Static R&
If R = 0 Then
    Dim Ws As Worksheet
    Set Ws = ActiveSheet
    R = Ws.Cells.Rows.Count
End If
MaxRow = R
End Function

Function NewWb(Optional Vis As Boolean) As Workbook
Dim O As Workbook
Set O = NewXls.Workbooks.Add
If Vis Then O.Application.Visible = True
Set NewWb = O
End Function

Function NewWs(Optional WsNm$, Optional Vis As Boolean) As Worksheet
Dim Wb As Workbook
Set Wb = NewWb
WsDlt Wb, "Sheet2"
WsDlt Wb, "Sheet3"
If WsNm <> "" Then WbWs(Wb, "Sheet1").Name = WsNm
Set NewWs = WbWs(Wb, 1)
If Vis Then WbVis Wb
End Function

Function NewWsA1() As Range
Set NewWsA1 = WsA1(NewWs)
End Function

Function NewXls() As Excel.Application
Static X As Excel.Application
On Error GoTo XX
Dim A$: A = X.Name
Set NewXls = X
Exit Function
XX:
Set X = New Excel.Application
Set NewXls = X
End Function

Function NmNxtSeqNm$(A, Optional NDig% = 3)
If NDig = 0 Then Stop
Dim R$: R = Right(A, NDig + 1)
If Left(R, 1) = "_" Then
    If IsNum(Mid(R, 2)) Then
        Dim L$: L = Left(A, Len(A) - NDig)
        Dim Nxt%: Nxt = Val(Mid(R, 2)) + 1
        NmNxtSeqNm = L + ZerFill(Nxt, NDig)
        Exit Function
    End If
End If
NmNxtSeqNm = A & "_" & StrDup(NDig - 1, "0") & "1"
End Function

Function NmSeqNo%(A)
Dim B$: B = TakAftRev(A, "_")
If B = "" Then Exit Function
If Not IsNum(B) Then Exit Function
NmSeqNo = B
End Function

Function RgA1(A As Range) As Range
Set RgA1 = RgRC(A, 1, 1)
End Function

Sub RgBdr(A As Range, Ix As XlBordersIndex, Optional Wgt As XlBorderWeight = xlMedium)
With A.Borders(Ix)
  .LineStyle = xlContinuous
  .Weight = Wgt
End With
End Sub

Sub RgBdrAround(A As Range)
A.BorderAround XlLineStyle.xlContinuous, xlMedium
If A.Row > 1 Then RgBdrBottom RgR(A, 0)
If A.Column > 1 Then RgBdrRight RgC(A, 0)
RgBdrTop RgR(A, RgNRow(A) + 1)
RgBdrLeft RgC(A, RgNCol(A) + 1)
End Sub

Sub RgBdrBottom(A As Range)
RgBdr A, xlEdgeBottom
End Sub

Sub RgBdrInside(A As Range)
RgBdr A, xlInsideHorizontal
RgBdr A, xlInsideVertical
End Sub

Sub RgBdrLeft(A As Range)
RgBdr A, xlEdgeLeft
If A.Column > 1 Then
    RgBdr RgC(A, A.Column - 1), xlEdgeRight
End If
End Sub

Sub RgBdrRight(A As Range)
RgBdr A, xlEdgeRight
If A.Column < MaxCol Then
    RgBdr RgC(A, A.Column + 1), xlEdgeLeft
End If
End Sub

Sub RgBdrTop(A As Range)
RgBdr A, xlEdgeTop
End Sub

Function RgC(A As Range, C) As Range
Set RgC = RgCRR(A, C, 1, A.Rows.Count)
End Function

Function RgCC(A As Range, C1, C2) As Range
Set RgCC = RgRCRC(A, 1, C1, RgNRow(A), C2)
End Function

Function RgCRR(A As Range, C, R1, R2) As Range
Set RgCRR = RgRCRC(A, R1, C, R2, C)
End Function

Function RgEntC(A As Range, C) As Range
Set RgEntC = RgC(A, C).EntireColumn
End Function

Function RgFstHBar(A As Range) As Range
Set RgFstHBar = RgR(A, 1)
End Function

Function RgFstVBar(A As Range) As Range
Set RgFstVBar = RgC(A, 1)
End Function

Function RgIsHBar(A As Range) As Boolean
RgIsHBar = A.Rows.Count = 1
End Function

Function RgIsVBar(A As Range) As Boolean
RgIsVBar = A.Columns.Count = 1
End Function

Function RgLasHBar(A As Range) As Range
Set RgLasHBar = RgR(A, RgNRow(A))
End Function

Function RgLasVBar(A As Range) As Range
Set RgLasVBar = RgC(A, RgNCol(A))
End Function

Sub RgLnkWs(A As Range)
Dim R As Range
Dim WsNy$(): WsNy = WbWsNy(RgWb(A))
For Each R In A
    CellLnkWs R, WsNy
Next
End Sub

Function RgLo(A As Range, Optional LoNm0$, Optional HasHeader As XlYesNoGuess = xlYes) As ListObject
Dim Ws As Worksheet: Set Ws = RgWs(A)
Dim O As ListObject: Set O = Ws.ListObjects.Add(xlSrcRange, A, , HasHeader)
If LoNm0 <> "" Then
    O.Name = WsDftLoNm(Ws, LoNm0)
End If
RgBdrAround A
Set RgLo = O
End Function

Function RgNCol%(A As Range)
RgNCol = A.Columns.Count
End Function

Function RgNRow%(A As Range)
RgNRow = A.Rows.Count
End Function

Function RgR(A As Range, R) As Range
Set RgR = RgRCC(A, R, 1, RgNCol(A))
End Function

Function RgRC(A As Range, R, C) As Range
Set RgRC = A.Cells(R, C)
End Function

Function RgRCC(A As Range, R, C1, C2) As Range
Set RgRCC = RgRCRC(A, R, C1, R, C2)
End Function

Function RgRCRC(Rg As Range, R1, C1, R2, C2) As Range
Dim Ws As Worksheet, Cell1 As Range, Cell2 As Range
Set Ws = Rg.Parent
Set Cell1 = RgRC(Rg, R1, C1)
Set Cell2 = RgRC(Rg, R2, C2)
Set RgRCRC = Ws.Range(Cell1, Cell2)
End Function

Function RgRR(A As Range, R1, R2) As Range
Set RgRR = RgRCRC(A, R1, 1, R2, RgNCol(A))
End Function

Function RgReSz(A As Range, Sq) As Range
Set RgReSz = RgRCRC(A, 1, 1, UBound(Sq, 1), UBound(Sq, 2))
End Function

Function RgSq(A As Range)
If A.Columns.Count = 1 Then
    If A.Rows.Count = 1 Then
        Dim O()
        ReDim O(1 To 1, 1 To 1)
        O(1, 1) = A.Value
        RgSq = O
        Exit Function
    End If
End If
RgSq = A.Value
End Function

Sub RgVis(A As Range)
WsVis RgWs(A)
End Sub

Function RgWb(A As Range) As Workbook
Set RgWb = WsWb(RgWs(A))
End Function

Function RgWs(A As Range) As Worksheet
Set RgWs = A.Parent
End Function

Sub RgeMgeV(A As Range)
Stop '?
End Sub

Function SampleDrs() As Drs
With SampleDrs
    .Fny = SampleDrsFny
    .Dry = SampleDry
End With
End Function

Function SampleDrsFny() As String()
SampleDrsFny = LvsSy("A B C D")
End Function

Function SampleDry() As Variant()
SampleDry = SqDry(SampleSq)
End Function

Function SampleLo() As ListObject
Set SampleLo = WsLo(SampleWs)
End Function

Function SampleSq()
Dim O()
ReDim O(1 To 10, 1 To 4)
Dim J%, I%
For J = 1 To 4
    For I = 1 To 10
        O(I, J) = I * 10 + J
    Next
Next
SampleSq = O
End Function

Function SampleWs() As Worksheet
Dim O As Worksheet
Set O = NewWs
DrsLo SampleDrs, WsA1(O)
Set SampleWs = O
WsVis O
End Function

Sub SqBrw(A)
DryBrw SqDry(A)
End Sub

Function SqCol(A, C%) As Variant()
Dim O()
Dim NR&, J&
NR = UBound(A, 1)
ReDim O(NR - 1)
For J = 1 To NR
    O(J - 1) = A(J, C)
Next
SqCol = O
End Function

Function SqDr(A, R&, Optional CnoAy) As Variant()
Dim mCnoAy%()
   Dim J%
   If IsMissing(CnoAy) Then
       ReDim mCnoAy(UBound(A, 2) - 1)
       For J = 0 To UB(mCnoAy)
           mCnoAy(J) = J + 1
       Next
   Else
       mCnoAy = CnoAy
   End If
Dim UCol%
   UCol = UB(mCnoAy)
Dim O()
   ReDim O(UCol)
   Dim C%
   For J = 0 To UCol
       C = mCnoAy(J)
       O(J) = A(R, C)
   Next
SqDr = O
End Function

Function SqDry(A) As Variant
Dim O(), NR&, NC%, R&, C%, UR&, UC%
NR = UBound(A, 1)
NC = UBound(A, 2)
UR = NR - 1
UC = NC - 1
Dim Dr()
For R = 1 To NR
    ReDim Dr(UC)
    For C = 1 To NC
        Dr(C - 1) = A(R, C)
    Next
    Push O, Dr
Next
SqDry = O
End Function

Function SqIsEmp(Sq) As Boolean
SqIsEmp = True
On Error GoTo X
Dim A
If UBound(Sq, 1) < 0 Then Exit Function
If UBound(Sq, 2) < 0 Then Exit Function
SqIsEmp = False
Exit Function
X:
End Function

Function SqRg(Sq, At As Range, Optional LoNm$) As Range
If SqIsEmp(Sq) Then Exit Function
Dim O As Range
Set O = RgReSz(At, Sq)
O.Value = Sq
Set SqRg = O
End Function

Function SqSel(Sq, Optional MapStr$) As Drs
Dim Fny$(), Fm$() 'MapStr
   If MapStr = "" Then
       Fny = AySy(SqDr(Sq, 1))
       Fm = Fny
   Else
       With S1S2AyStr_SyPair(MapStr)
           Fny = .Sy1
           Fm = .Sy2
       End With
   End If
Dim SqCnoAy%() 'Fm,Sq
   Dim A&()
   Dim U%
   Dim J%
   A = AyIxAy(SqDr(Sq, 1), Fm)
   U = UB(A)
   ReDim SqCnoAy(U)
   For J = 0 To U
       SqCnoAy(J) = A(J) + 1
   Next
Dim Dry() 'Sq,SqIxAy
   Dim R&, Cno%, C%
   Dim UFld%
   Dim Ix%
   Dim Dr()
   UFld = UB(SqCnoAy)
   For R = 2 To UBound(Sq, 1)
       ReDim Dr(UFld)
       For C = 0 To UFld
           Cno = SqCnoAy(C)
           If Cno > 0 Then
               Dr(C) = Sq(R, Cno)
           End If
       Next
       Push Dry, Dr
   Next
SqSel.Dry = Dry
SqSel.Fny = Fny
End Function

Sub TitRg_Fmt(A As Range)
'---
    Dim C%
    For C = 1 To A.Columns.Count
        VBar_MgeBottomEmpCell RgC(A, C)
    Next
HBar_MgeSamValCell A
End Sub

Function TitS1S2Ay_Sq(TitS1S2Ay() As S1S2, Fny$()) As Variant()
Dim TitColAy$():   TitColAy = A_TitColAy(TitS1S2Ay, Fny)
Dim VBarColAy():  VBarColAy = A_VBarColAy(TitColAy)
Dim NRow%:             NRow = A_TitNRow(VBarColAy)
Dim Sq(): ReDim Sq(1 To NRow, 1 To Sz(Fny))
Dim J%, C%, R%, VBar$()
For C = 0 To UB(Fny)
    VBar = VBarColAy(C)
    For R = 0 To UB(VBar)
        Sq(R + 1, C + 1) = VBar(R)
    Next
Next
TitS1S2Ay_Sq = Sq
End Function

Function VBarAy(A As Range) As Variant()
Ass RgIsVBar(A)
VBarAy = SqCol(RgSq(A), 1)
End Function

Function VBarIntAy(A As Range) As Integer()
VBarIntAy = AyIntAy(VBarAy(A))
End Function

Function VBarSy(A As Range) As String()
VBarSy = AySy(VBarAy(A))
End Function

Sub VBar_MgeBottomEmpCell(A As Range)
Ass RgIsVBar(A)
Dim R2: R2 = A.Rows.Count
Dim R1
    Dim Fnd As Boolean
    For R1 = R2 To 1 Step -1
        If Not IsEmpty(RgRC(A, R1, 1)) Then Fnd = True: GoTo Nxt
    Next
Nxt:
    If Not Fnd Then Stop
If R2 = R1 Then Exit Sub
Dim R As Range: Set R = RgCRR(A, 1, R1, R2)
R.Merge
R.VerticalAlignment = XlVAlign.xlVAlignTop
End Sub

Function WbAddDt(A As Workbook, Dt As Dt) As Worksheet
Dim O As Worksheet
Set O = WbAddWs(A, Dt.DtNm)
DrsLo DtDrs(Dt), WsA1(O), Dt.DtNm
Set WbAddDt = O
End Function

Function WbAddWs(A As Workbook, Optional WsNm$, Optional IsBeg As Boolean) As Worksheet
Dim O As Worksheet
If IsBeg Then
    Set O = A.Sheets.Add(WbFstWs(A))
Else
    Set O = A.Sheets.Add(, WbLasWs(A))
End If
If WsNm <> "" Then
   O.Name = WsNm
End If
Set WbAddWs = O
End Function

Sub WbClsNoSav(A As Workbook)
On Error Resume Next
A.Close False
End Sub

Function WbFstWs(A As Workbook) As Worksheet
Set WbFstWs = A.Sheets(1)
End Function

Function WbHasWs(A As Workbook, Ix_or_WsNm) As Boolean
On Error GoTo X
Dim Ws As Worksheet: Set Ws = A.Sheets(Ix_or_WsNm)
WbHasWs = True
Exit Function
X:
End Function

Function WbLasWs(A As Workbook) As Worksheet
Set WbLasWs = A.Sheets(A.Sheets.Count)
End Function

Sub WbSav(A As Workbook)
Dim X As Excel.Application
Set X = A.Application
Dim Y As Boolean
Y = X.DisplayAlerts
A.DisplayAlerts = False
A.Save
A.DisplayAlerts = Y
End Sub

Sub WbVis(A As Workbook)
A.Application.Visible = True
End Sub

Function WbWs(A As Workbook, Ix_or_WsNm) As Worksheet
Set WbWs = A.Sheets(Ix_or_WsNm)
End Function

Function WbWsNy(A As Workbook) As String()
WbWsNy = ItrNy(A.Sheets)
End Function

Function WsA1(A As Worksheet) As Range
Set WsA1 = A.Range("A1")
End Function

Function WsCRR(A As Worksheet, C, R1, R2) As Range
Set WsCRR = WsRCRC(A, R1, C, R2, C)
End Function

Sub WsClsNoSav(A As Worksheet)
WbClsNoSav WsWb(A)
End Sub

Function WsDftLoNm$(A As Worksheet, Optional LoNm0$)
Dim LoNm$: LoNm = DftStr(LoNm0, "Table")
Dim J%
For J = 1 To 999
    If Not WsHasLoNm(A, LoNm) Then WsDftLoNm = LoNm: Exit Function
    LoNm = NmNxtSeqNm(LoNm)
Next
Stop
End Function

Function WsDlt(A As Workbook, WsIx) As Boolean
If WbHasWs(A, WsIx) Then WbWs(A, WsIx).Delete
WsDlt = True
End Function

Function WsDtaRg(A As Worksheet) As Range
Dim R, C
With WsLasCell(A)
   R = .Row
   C = .Column
End With
If R = 1 And C = 1 Then Exit Function
Set WsDtaRg = WsRCRC(A, 1, 1, R, C)
End Function

Function WsHasLoNm(A As Worksheet, LoNm$) As Boolean
Dim I
For Each I In A.ListObjects
    If CvLo(I).Name = LoNm Then WsHasLoNm = True: Exit Function
Next
End Function

Function WsLasCell(A As Worksheet) As Range
Set WsLasCell = A.Cells.SpecialCells(xlCellTypeLastCell)
End Function

Function WsLasCno%(A As Worksheet)
WsLasCno = WsLasCell(A).Column
End Function

Function WsLasRno%(A As Worksheet)
WsLasRno = WsLasCell(A).Row
End Function

Function WsLo(A As Worksheet, Optional LoNm$ = "Table1") As ListObject
Dim O As ListObject
Set O = A.ListObjects.Add(xlSrcRange, WsDtaRg(A), xlNo)
Dim N$: N = WsDftLoNm(A, LoNm)
If LoNm <> N Then O.Name = N
Set WsLo = O
End Function

Function WsNxtLoNm$(A As Worksheet, LoNm$)

End Function

Function WsRC(A As Worksheet, R, C) As Range
Set WsRC = A.Cells(R, C)
End Function

Function WsRCRC(A As Worksheet, R1, C1, R2, C2) As Range
Set WsRCRC = A.Range(WsRC(A, R1, C1), WsRC(A, R2, C2))
End Function

Function WsSq(A As Worksheet) As Variant()
WsSq = WsDtaRg(A).Value
End Function

Sub WsVis(A As Worksheet)
A.Application.Visible = True
End Sub

Function WsWb(A As Worksheet) As Workbook
Set WsWb = A.Parent
End Function

Function ZerFill$(N%, NDig%)
ZerFill = Format(N, StrDup(NDig, 0))
End Function

Private Function A_TitColAy(TitS1S2Ay() As S1S2, Fny$()) As String()
Dim O$(), J%, I%, UTit%
UTit = S1S2_UB(TitS1S2Ay)
For J = 0 To UB(Fny)
    For I = 0 To UTit
        If TitS1S2Ay(I).S1 = Fny(J) Then Push O, TitS1S2Ay(I).S2: GoTo Nxt
    Next
    Push O, Fny(J)
Nxt:
Next
A_TitColAy = O
End Function

Private Function A_TitNRow%(VBarColAy())
Dim M%, J%, S%
For J = 0 To UB(VBarColAy)
    S = Sz(VBarColAy(J))
    If S > M Then M = S
Next
A_TitNRow = M
End Function

Private Function A_VBarColAy(TitColAy$()) As Variant()
Dim O(), J%
For J = 0 To UB(TitColAy)
    Dim VBar$()
    VBar = AyTrim(SplitVBar(TitColAy(J)))
    Push O, VBar
Next
A_VBarColAy = O
End Function

Private Sub FxAqlDrs__Tst()
Const Fx$ = "N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls"
Const Sql$ = "Select * from [Sheet1$]"
DrsBrw FxSql_Drs(Fx, Sql)
End Sub

Private Sub FxCat__Tst()
Dim A As ADOX.Catalog
Set A = FxCat("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
Stop
End Sub

Private Sub FxCn__Tst()
Dim A As ADODB.Connection
Set A = FxCn("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
Stop
End Sub

Private Sub FxRmvWsIfExist__Tst()
Dim T$: T = TmpFx("FxRmvWsIfExist__Tst")
Dim Wb As Workbook
Set Wb = NewWb
Wb.Sheets.Add
Wb.SaveAs T
Dim WsNyBef$(), WsNyAft$()
   WsNyBef = FxWsNy(T)
   FxRmvWsIfExist T, "Sheet1"
   WsNyAft = FxWsNy(T)
Dim Exp$()
   Exp = AyMinus(WsNyBef, Array("Sheet1"))
AyPair_EqChk Exp, WsNyAft
End Sub

Private Sub FxSql_Arun__Tst()
Const Fx$ = "N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls"
Const Sql$ = "Select * into [Sheet21] from [Sheet1$]"
FxSql_Arun Fx, Sql
End Sub

Private Sub FxWsFny__Tst()
AyDmp FxWsFny("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls", "Sheet1")
End Sub

Private Sub FxWsNy__Tst()
AyDmp FxWsNy("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
End Sub

Sub FxaCrt__Tst()
Dim Act As Excel.Application
Set Act = FxaCrt
Act.Visible = True
Stop
End Sub

Sub LoAdjColWdt__Tst()
Dim Ws As Worksheet: Set Ws = NewWs(Vis:=True)
Dim Sq(1 To 2, 1 To 2)
Sq(1, 1) = "A"
Sq(1, 2) = "B"
Sq(2, 1) = "123123"
Sq(2, 2) = String(1234, "A")
Ws.Range("A1:B2").Value = Sq
LoAdjColWdt LoCrt(Ws)
WsClsNoSav Ws
End Sub

Sub LoBrw__Tst()
Dim O As ListObject: Set O = SampleLo
LoBrw O
Stop
End Sub

Private Sub TitS1S2Ay_Sq__Tst()
Dim Fny$()
    PushAy Fny, Array("X", "A", "C", "B")
Dim TitS1S2Ay() As S1S2
    S1S2_Push TitS1S2Ay, NewS1S2("A", "skldf|lsjdf|sdldf")
    S1S2_Push TitS1S2Ay, NewS1S2("C", "skldf|lsjdf|sdlkf |sdfsdf")
    S1S2_Push TitS1S2Ay, NewS1S2("B", "skldf|ls|df|jdf|sdlkf |sdfsdf")
    S1S2_Push TitS1S2Ay, NewS1S2("X", "skdf|df|lsjdf|sdlkf |sdfsdf")
SqBrw TitS1S2Ay_Sq(TitS1S2Ay, Fny)
End Sub


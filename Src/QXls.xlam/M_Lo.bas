Attribute VB_Name = "M_Lo"
Option Explicit

Property Get LoC(A As ListObject, C, Optional InclTot As Boolean, Optional InclHdr As Boolean) As Range
Dim R As Range
Set R = A.ListColumns(C).DataBodyRange
If Not InclTot And Not InclHdr Then
    Set LoC = R
    Exit Property
End If

Dim R1&, R2&
    R1 = 1
    R2 = R.Rows.Count
    If InclTot Then R2 = R2 + 1
    If InclHdr Then R1 = R1 - 1
Set LoC = RgRR(R, R1, R2)
End Property

Property Get LoCC(A As ListObject, C1, C2, Optional InclTot As Boolean, Optional InclHdr As Boolean) As Range
Dim R1&, R2&, mC1%, mC2%
R1 = LoR1(A, InclHdr)
R2 = LoR2(A, InclTot)
mC1 = LoWsCno(A, C1)
mC2 = LoWsCno(A, C2)
Set LoCC = WsRCRC(LoWs(A), R1, mC1, R2, mC2)
End Property

Property Get LoCol_Rg(A As ListObject, ColNm$) As Range
Set LoCol_Rg = A.ListColumns(ColNm).Range
End Property

Property Get LoCrt(A As Worksheet, Optional LoNm$) As ListObject
Dim R As Range: Set R = WsDtaRg(A)
If IsNothing(R) Then Exit Property
Dim O As ListObject: Set O = A.ListObjects.Add(xlSrcRange, WsDtaRg(A), , xlYes)
If LoNm <> "" Then O.Name = LoNm
LoAdjColWdt O
Set LoCrt = O
End Property

Property Get LoDry(A As ListObject) As Variant()
LoDry = SqDry(LoSq(A))
End Property

Property Get LoEntCol(A As ListObject) As Range
Set LoEntCol = LoCC(A, 1, LoNCol(A)).EntireColumn
End Property

Property Get LoFny(A As ListObject) As String()
Dim O$(), I
For Each I In A.ListColumns
    Push O, CvLoCol(I).Name
Next
LoFny = O
End Property

Property Get LoHasNoDta(A As ListObject) As Boolean
LoHasNoDta = IsNothing(A.DataBodyRange)
End Property

Property Get LoHdrCell(A As ListObject, FldNm) As Range
Dim Rg As Range: Set Rg = A.ListColumns(FldNm).Range
Set LoHdrCell = RgRC(Rg, 1, 1)
End Property

Property Get LoNCol%(A As ListObject)
LoNCol = A.ListColumns.Count
End Property

Property Get LoR1&(A As ListObject, Optional InclHdr As Boolean)
If LoHasNoDta(A) Then
   LoR1 = A.ListColumns(1).Range.Row + 1
   Exit Property
End If
LoR1 = A.DataBodyRange.Row - IIf(InclHdr, 1, 0)
End Property

Property Get LoR2&(A As ListObject, Optional InclTot As Boolean)
If LoHasNoDta(A) Then
   LoR2 = LoR1(A)
   Exit Property
End If
LoR2 = A.DataBodyRange.Row + IIf(InclTot, 1, 0)
End Property

Property Get LoSq(A As ListObject)
LoSq = A.DataBodyRange.Value
End Property

Property Get LoWs(A As ListObject) As Worksheet
Set LoWs = A.Parent
End Property

Property Get LoWsCno%(A As ListObject, Ix_or_ColNm)
LoWsCno = A.ListColumns(Ix_or_ColNm).Range.Column
End Property

Sub LoAdjColWdt(A As ListObject)
Dim C As Range: Set C = LoEntCol(A)
C.AutoFit
Dim EntC As Range, J%
For J = 1 To C.Columns.Count
   Set EntC = RgEntC(C, J)
   If EntC.ColumnWidth > 100 Then EntC.ColumnWidth = 100
Next
End Sub

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

Sub LoVis(A As ListObject)
A.Application.Visible = True
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
'LoBrw O
Stop
End Sub

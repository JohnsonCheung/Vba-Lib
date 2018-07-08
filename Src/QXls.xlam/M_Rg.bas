Attribute VB_Name = "M_Rg"
Option Explicit

Property Get RgA1(A As Range) As Range
Set RgA1 = RgRC(A, 1, 1)
End Property

Property Get RgC(A As Range, C) As Range
Set RgC = RgCRR(A, C, 1, A.Rows.Count)
End Property

Property Get RgCC(A As Range, C1, C2) As Range
Set RgCC = RgRCRC(A, 1, C1, RgNRow(A), C2)
End Property

Property Get RgCRR(A As Range, C, R1, R2) As Range
Set RgCRR = RgRCRC(A, R1, C, R2, C)
End Property

Property Get RgEntC(A As Range, C) As Range
Set RgEntC = RgC(A, C).EntireColumn
End Property

Property Get RgFstHBar(A As Range) As Range
Set RgFstHBar = RgR(A, 1)
End Property

Property Get RgFstVBar(A As Range) As Range
Set RgFstVBar = RgC(A, 1)
End Property

Property Get RgIsHBar(A As Range) As Boolean
RgIsHBar = A.Rows.Count = 1
End Property

Property Get RgIsVBar(A As Range) As Boolean
RgIsVBar = A.Columns.Count = 1
End Property

Property Get RgLasHBar(A As Range) As Range
Set RgLasHBar = RgR(A, RgNRow(A))
End Property

Property Get RgLasVBar(A As Range) As Range
Set RgLasVBar = RgC(A, RgNCol(A))
End Property

Property Get RgLo(A As Range, Optional LoNm0$, Optional HasHeader As XlYesNoGuess = xlYes) As ListObject
Dim Ws As Worksheet: Set Ws = RgWs(A)
Dim O As ListObject: Set O = Ws.ListObjects.Add(xlSrcRange, A, , HasHeader)
If LoNm0 <> "" Then
    O.Name = WsDftLoNm(Ws, LoNm0)
End If
RgBdrAround A
Set RgLo = O
End Property

Property Get RgNCol%(A As Range)
RgNCol = A.Columns.Count
End Property

Property Get RgNRow%(A As Range)
RgNRow = A.Rows.Count
End Property

Property Get RgR(A As Range, R) As Range
Set RgR = RgRCC(A, R, 1, RgNCol(A))
End Property

Property Get RgRC(A As Range, R, C) As Range
Set RgRC = A.Cells(R, C)
End Property

Property Get RgRCC(A As Range, R, C1, C2) As Range
Set RgRCC = RgRCRC(A, R, C1, R, C2)
End Property

Property Get RgRCRC(Rg As Range, R1, C1, R2, C2) As Range
Dim Ws As Worksheet, Cell1 As Range, Cell2 As Range
Set Ws = Rg.Parent
Set Cell1 = RgRC(Rg, R1, C1)
Set Cell2 = RgRC(Rg, R2, C2)
Set RgRCRC = Ws.Range(Cell1, Cell2)
End Property

Property Get RgRR(A As Range, R1, R2) As Range
Set RgRR = RgRCRC(A, R1, 1, R2, RgNCol(A))
End Property

Property Get RgReSz(A As Range, Sq) As Range
Set RgReSz = RgRCRC(A, 1, 1, UBound(Sq, 1), UBound(Sq, 2))
End Property

Property Get RgSq(A As Range)
If A.Columns.Count = 1 Then
    If A.Rows.Count = 1 Then
        Dim O()
        ReDim O(1 To 1, 1 To 1)
        O(1, 1) = A.Value
        RgSq = O
        Exit Property
    End If
End If
RgSq = A.Value
End Property

Property Get RgWb(A As Range) As Workbook
Set RgWb = WsWb(RgWs(A))
End Property

Property Get RgWs(A As Range) As Worksheet
Set RgWs = A.Parent
End Property

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
    RgBdr RgC(A, 0), xlEdgeRight
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

Sub RgLnkWs(A As Range)
Dim R As Range
Dim WsNy$(): WsNy = WbWsNy(RgWb(A))
For Each R In A
    CellLnkWs R, WsNy
Next
End Sub

Sub RgVis(A As Range)
WsVis RgWs(A)
End Sub

Sub RgeMgeV(A As Range)
Stop '?
End Sub

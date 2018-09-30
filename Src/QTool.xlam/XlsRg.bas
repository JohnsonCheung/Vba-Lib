Attribute VB_Name = "XlsRg"
Option Explicit
Function RgIncTopR(A As Range, Optional By% = 1) As Range
Set RgIncTopR = RgRR(A, 1 - By, A.Rows.Count)
End Function
Function RgRR(A As Range, R1, R2) As Range
Set RgRR = RgRCRC(A, R1, 1, R2, A.Columns.Count)
End Function
Function RgLo(A As Range, Optional LoNm$) As ListObject
Dim O As ListObject
Set O = RgWs(A).ListObjects.Add(xlSrcRange, A, , xlYes)
If LoNm <> "" Then O.Name = LoNm
Set RgLo = O
End Function
Function RgRC(A As Range, R, C) As Range
Set RgRC = A.Cells(R, C)
End Function
Function RgRCRC(A As Range, R1, C1, R2, C2) As Range
Set RgRCRC = RgWs(A).Range(RgRC(A, R1, C1), RgRC(A, R2, C2))
End Function
Function RgWs(A As Range)
Set RgWs = A.Parent
End Function
Sub RgBdrTop(A As Range)
RgBdr A, xlEdgeTop
End Sub
Sub RgBdr(A As Range, Ix As XlBordersIndex, Optional Wgt As XlBorderWeight = xlMedium)
With A.Borders(Ix)
  .LineStyle = xlContinuous
  .Weight = Wgt
End With
End Sub
Function RgR(A As Range, R) As Range
Set RgR = RgRCRC(A, R, 1, R, RgNCol(A))
End Function
Function RgC(A As Range, C) As Range
Set RgC = RgRCRC(A, 1, RgNRow(A), 1, C)
End Function
Function RgNRow&(A As Range)
RgNRow = A.Rows.Count
End Function
Function RgNCol%(A As Range)
RgNCol = A.Columns.Count
End Function
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
Sub RgVis(A As Range, Vis As Boolean)
If Vis Then A.Application.Visible = True
End Sub
Function RgOf_FmPj(A As Worksheet) As Range
Set RgOf_FmPj = WsRCRC(A, 1, 1, 3, 1)
End Function
Function RgOf_ToPj(A As Worksheet) As Range
Set RgOf_ToPj = WsRCRC(A, 1, 2, 3, 2)
End Function

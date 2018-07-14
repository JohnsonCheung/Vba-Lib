Attribute VB_Name = "M_Cell"
Option Explicit
Property Get CellAyH(A As Range) As Variant()
If IsEmpty(RgA1(A).Value) Then Exit Property
Dim R&
For R = 2 To MaxRow
    If IsEmpty(RgRC(A, R, 1).Value) Then
        CellAyH = SqCol(RgCRR(A, 1, 1, R - 1).Value, 1)
        Exit Property
    End If
Next
Stop
End Property

Property Get CellAyV(A As Range) As Variant()

End Property

Property Get CellIsInRg(A As Range, Rg As Range) As Boolean
Dim R&, C%, R1&, R2&, C1%, C2%
R = A.Row
R1 = Rg.Row
If R < R1 Then Exit Property
R2 = R1 + Rg.Rows.Count
If R > R2 Then Exit Property
C = A.Column
C1 = Rg.Column
If C < C1 Then Exit Property
C2 = C1 + Rg.Columns.Count
If C > C2 Then Exit Property
CellIsInRg = True
End Property

Property Get CellIsInRgAp(A As Range, ParamArray RgAp()) As Boolean
Dim Av(): Av = RgAp
CellIsInRgAp = CellIsInRgAv(A, Av)
End Property

Property Get CellIsInRgAv(A As Range, RgAv()) As Boolean
Dim V
For Each V In RgAv
    If CellIsInRg(A, CvRg(V)) Then CellIsInRgAv = True: Exit Property
Next
End Property

Property Get CellVBar(Cell As Range, Optional AtLeastOneCell As Boolean) As Range
If IsEmpty(Cell.Value) Then
    If AtLeastOneCell Then
        Set CellVBar = RgA1(Cell)
    End If
    Exit Property
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
End Property

Sub CellClrDown(A As Range)
CellVBar(A, AtLeastOneCell:=True).Clear
End Sub

Sub CellFillSeqDown(A As Range, N&, Optional IsFmOne As Boolean)
AyRgV UIntSeq(N, IsFmOne), A
End Sub

Sub CellLnkWs(A As Range, WsNy$())
Dim WsNm: WsNm = A.Value
If Not IsStr(WsNm) Then Exit Sub
If Not AyHas(WsNy, WsNm) Then Exit Sub
With A.Hyperlinks
    If .Count > 0 Then .Delete
    .Add A, "", FmtQQ("'?'!A1", WsNm)
End With
End Sub

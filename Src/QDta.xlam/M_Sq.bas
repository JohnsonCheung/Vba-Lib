Attribute VB_Name = "M_Sq"
Option Explicit

Function SqNCol&(A)
On Error Resume Next
SqNCol = UBound(A, 2)
End Function

Function SqNRow&(A)
On Error Resume Next
SqNRow = UBound(A, 1)
End Function

Function SqRg(A, At As Range, Optional LoNm$) As Range
If Sz(A) = 0 Then Set SqRg = At.Cells(1, 1): Exit Function
Dim O As Range
Set O = RgReSz(At, A)
O.Value = A
Set SqRg = O
End Function

Function SqTranspose(A) As Variant()
Dim NRow&, NCol&
NRow = SqNRow(A)
NCol = SqNCol(A)
Dim O(), J&, I&
ReDim O(1 - NCol, 1 To NRow)
For J = 1 To NRow
    For I = 1 To NCol
        O(I, J) = A(J, I)
    Next
Next
SqTranspose = O
End Function

Sub SqBrw(A)
DryBrw SqDry(A)
End Sub

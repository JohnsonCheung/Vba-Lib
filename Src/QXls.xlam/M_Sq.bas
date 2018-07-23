Attribute VB_Name = "M_Sq"
Option Explicit

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


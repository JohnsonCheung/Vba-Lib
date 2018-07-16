Attribute VB_Name = "M_Sq"
Option Explicit

Property Get SqCol(A, C%) As Variant()
Dim O()
Dim NR&, J&
NR = UBound(A, 1)
ReDim O(NR - 1)
For J = 1 To NR
    O(J - 1) = A(J, C)
Next
SqCol = O
End Property

Property Get SqDr(A, R&, Optional CnoAy) As Variant()
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
End Property



Property Get SqDry(A) As Variant
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
End Property

Property Get SqIsEmp(Sq) As Boolean
SqIsEmp = True
On Error GoTo X
Dim A
If UBound(Sq, 1) < 0 Then Exit Property
If UBound(Sq, 2) < 0 Then Exit Property
SqIsEmp = False
Exit Property
X:
End Property

Property Get SqRg(Sq, At As Range, Optional LoNm$) As Range
If SqIsEmp(Sq) Then Exit Property
Dim O As Range
Set O = RgReSz(At, Sq)
O.Value = Sq
Set SqRg = O
End Property


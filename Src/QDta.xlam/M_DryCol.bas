Attribute VB_Name = "M_DryCol"
Option Explicit

Function DryCol_Into(A, ColIx%, OIntoAy)
Dim O: O = OIntoAy: Erase O
If Sz(A) = 0 Then
    DryCol_Into = O
    Exit Function
End If
Dim Dr, J&
ReDim O(UB(A))
For Each Dr In A
    If UB(Dr) >= ColIx Then
        O(J) = Dr(ColIx)
    End If
    J = J + 1
Next
End Function

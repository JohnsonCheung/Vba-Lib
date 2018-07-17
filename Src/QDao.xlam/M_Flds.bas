Attribute VB_Name = "M_Flds"
Option Explicit
Function FldsDr(A As Dao.Fields) As Variant()
Dim O(), J%
ReDim O(A.Count - 1)
For J = 0 To A.Count - 1
   O(J) = A(J).Value
Next
FldsDr = O
End Function
Function FldsFny(A As Dao.Fields) As String()
Dim O$(), J%
ReDim O(A.Count - 1)
For J = 0 To A.Count - 1
   O(J) = A(J).Name
Next
FldsFny = O
End Function
Function FldsHasFld(A As Dao.Fields, F) As Boolean
Dim I  As Dao.Field
For Each I In A
   If I.Name = F Then FldsHasFld = True: Exit Function
Next
End Function

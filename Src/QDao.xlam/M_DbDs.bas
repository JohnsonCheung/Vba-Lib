Attribute VB_Name = "M_DbDs"
Option Explicit
Function DbDs_SqlAy_OfIns(A As Database, Ds As Ds) As String()
If DsIsEmp(Ds) Then Exit Function
Dim O$()
Dim J%
For J = 0 To UBound(Ds.DtAy)

   PushAy O, DbDt_SqlAy_OfIns(A, DsDt(Ds, J))
Next
DbDs_SqlAy_OfIns = O
End Function
Function DsDt(A As Ds, Ix%) As Dt
Dim DtAy() As Dt
DtAy = A.DtAy
Set DsDt = DtAy(Ix)
End Function

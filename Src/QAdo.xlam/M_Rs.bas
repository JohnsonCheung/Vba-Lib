Attribute VB_Name = "M_Rs"
Option Explicit

Function RsDrs(A As Recordset) As Drs
Set RsDrs = Drs(RsFny(A), RsDry(A))
End Function

Function RsDry(A As Recordset) As Variant()
Dim O()
While Not A.EOF
    Push O, FldsDr(A.Fields)
    A.MoveNext
Wend
RsDry = O
End Function

Function RsFny(A As Recordset) As String()
RsFny = FldsFny(A.Fields)
End Function


Attribute VB_Name = "M_Rs"
Option Explicit

Property Get RsDrs(A As Recordset) As Drs
Set RsDrs = Drs(RsFny(A), RsDry(A))
End Property

Property Get RsDry(A As Recordset) As Variant()
Dim O()
While Not A.EOF
    Push O, FldsDr(A.Fields)
    A.MoveNext
Wend
RsDry = O
End Property

Property Get RsFny(A As Recordset) As String()
RsFny = FldsFny(A.Fields)
End Property


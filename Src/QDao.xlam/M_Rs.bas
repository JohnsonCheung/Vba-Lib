Attribute VB_Name = "M_Rs"
Option Explicit

Function RsFny(A As Recordset) As String()
RsFny = FldsFny(A.Fields)
End Function

Function RsSy(A As Recordset) As String()
Dim O$()
With A
   While Not .EOF
       Push O$, A.Fields(0).Value
       .MoveNext
   Wend
End With
RsSy = O
End Function

Function RsDry(A As Recordset) As Variant()
Dim O()
With A
    While Not .EOF
        Push O, FldsDr(A.Fields)
        .MoveNext
    Wend
End With
RsDry = O
End Function

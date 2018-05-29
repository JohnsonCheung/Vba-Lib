Attribute VB_Name = "Ado"
Option Explicit
Property Get ARs(A As ADODB.Recordset) As ARs
Dim O As New ARs
Set O.Rs = A
Set ARs = O
End Property

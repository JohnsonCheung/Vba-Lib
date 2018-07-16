Attribute VB_Name = "JAdo"
Option Explicit


Property Get Tst() As Tst
Static A As New Tst
Set Tst = A
End Property
Property Get AFb(A$) As AFb
Dim O As New AFb
Set AFb = O.Init(A)
End Property
Property Get ARs(A As Recordset) As ARs
Dim O As New ARs
Set ARs = O.Init(A)
End Property
Property Get ACn(A As Connection) As ACn
Dim O As New ACn
Set ACn = O.Init(A)
End Property
Property Get AFlds(A As Fields) As AFlds
Dim O As New AFlds
Set AFlds = O.Init(A)
End Property
Property Get AFx(A) As AFx
Dim O As New AFx
Set AFx = O.Init(A)
End Property

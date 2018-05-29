Attribute VB_Name = "Ado"
Option Explicit

Property Get AdoP() As AdoPrivateClassCreator
Static Y As New AdoPrivateClassCreator
Set AdoP = Y
End Property

Property Get Tst() As AdoTst
Static A As New AdoTst
Set Tst = A
End Property
Property Get Fb(A$) As Fb
Dim O As New Fb
Set Fb = O.Init(A)
End Property

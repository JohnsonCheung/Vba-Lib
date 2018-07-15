Attribute VB_Name = "JIde"
Option Explicit

Property Get Mth(A As CodeModule, Nm$) As Mth
Dim O As New Mth
Set Mth = O.Init(A, Nm)
End Property

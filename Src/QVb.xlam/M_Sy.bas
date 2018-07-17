Attribute VB_Name = "M_Sy"
Option Explicit
Property Get Sy_OfBoolOp() As String()
Static Y$(), X As Boolean
If Not X Then
    X = True
    Y = SslSy("AND OR")
End If
Sy_OfBoolOp = Y
End Property

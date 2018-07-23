Attribute VB_Name = "M_SyOf"
Option Explicit

Function SyOf_BoolOp() As String()
Static Y$(), X As Boolean
If Not X Then
    X = True
    Y = SslSy("AND OR")
End If
SyOf_BoolOp = Y
End Function




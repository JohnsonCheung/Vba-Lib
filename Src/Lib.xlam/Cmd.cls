VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Property Get Lis() As CmdLis
Static Y As New CmdLis
Set Lis = Y
End Property

Function IEq_Eq(A As Variant, B As Variant) As Boolean

End Function

Function IEq_Ne(A As Variant, B As Variant) As Boolean

End Function

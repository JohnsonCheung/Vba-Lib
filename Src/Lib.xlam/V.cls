VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "V"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private A_V
Sub Init(A)
A_V = A
End Sub
Function IsStr() As Boolean
IsStr = ValIsStr(A_V)
End Function

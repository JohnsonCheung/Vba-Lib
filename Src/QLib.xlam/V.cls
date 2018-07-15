VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "V"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_V

Sub Init(A)
A_V = A
End Sub

Function IsStr() As Boolean
IsStr = VarIsStr(A_V)
End Function

Function IsSy() As Boolean
IsSy = Ty = vbArray + vbString
End Function

Function Ty() As VbVarType
Ty = VarType(A_V)
End Function

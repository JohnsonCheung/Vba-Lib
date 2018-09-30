VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WhNm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Re As RegExp
Dim X_ExlAy$()
Property Get ExlAy() As String()
ExlAy = X_ExlAy
End Property
Property Let ExlAy(V$())
X_ExlAy = V
End Property


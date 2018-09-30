VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WhMth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Nm As WhNm
Dim X_InShtMdy$(), X_InShtKd$()
Property Get InShtMdy() As String()
InShtMdy = X_InShtMdy
End Property
Property Get InShtKd() As String()
InShtKd = X_InShtKd
End Property
Property Let InShtMdy(V$())
X_InShtMdy = V
End Property
Property Let InShtKd(V$())
X_InShtKd = V
End Property


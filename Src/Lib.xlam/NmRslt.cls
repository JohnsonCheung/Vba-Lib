VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NmRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Nm$, Er As New Er
Property Get ToStr$()
ToStr = "Nm(?) " & Er.ToStr
End Property
Property Get Lin$()
Stop
End Property

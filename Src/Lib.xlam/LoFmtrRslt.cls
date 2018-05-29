VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoFmtrRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public LoFmtr As New LoFmtr, Er As New Er, FmtWs As New FmtWs
Property Get ToStr$()
ToStr = Oy(Array(LoFmtr, FmtWs, Er)).ToStr
End Property

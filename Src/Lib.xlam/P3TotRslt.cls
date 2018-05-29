VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "P3TotRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Sum As LABCsRslt
Public Avg As LABCsRslt
Public Cnt As LABCsRslt
Private B_ABCLy() As String
Public Er As New Er
Property Get ABCLy() As String()
ABCLy = B_ABCLy
End Property

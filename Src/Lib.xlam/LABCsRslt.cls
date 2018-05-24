VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LABCsRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public LABCs As LABCs
Public Er As Er
Function Init(A As LABCs, Optional Er As Er) As LABCsRslt
Set LABCs = A
If IsNothing(Er) Then Set Me.Er = New Er Else Set Me.Er = Er
Set Init = Me
End Function

'Public Er As New Er
Property Get ToStr$()
ToStr = StrAp_Lines(LABCs.ToStr, Er.ToStr)
End Property


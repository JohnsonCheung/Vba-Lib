VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LABCsRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public LABCs As New LABCs
Public Er As New Er

Function Init(A As LABCs, Optional Er As Er) As LABCsRslt
Set LABCs = A
If Not IsNothing(Er) Then Set Me.Er = Er
Set Init = Me
End Function

Property Get ToStr$()
ToStr = StrAp_Lines(LABCs.ToStr, Er.ToStr)
End Property


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FnyRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private B_Fny$()
Public Er As Er
Property Get Fny() As String()
Fny = B_Fny
End Property
Function Init(Fny$(), Er As Er) As FnyRslt
B_Fny = Fny
Set Me.Er = Er
Set Init = Me
End Function
Property Get Lin$()
Lin = C1_Lo & " " & C2_Lo_Fld & " " & JnSpc(B_Fny)
End Property

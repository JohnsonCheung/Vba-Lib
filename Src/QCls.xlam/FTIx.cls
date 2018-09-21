VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FTIx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Fmix%, Toix%

Friend Property Get Init(Fmix%, Toix%) As FTIx
If Fmix < 0 Then Stop
If Toix < 0 Then Stop
If Fmix > Toix - 1 Then Stop
Me.Fmix = Fmix
Me.Toix = Toix
Set Init = Me
End Property

Property Get ToStr$()
ToStr = "FTIx(Fm(" & Fmix & ") To(" & Toix & "))"
End Property

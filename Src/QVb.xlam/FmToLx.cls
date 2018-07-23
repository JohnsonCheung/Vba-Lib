VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FmToLx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public FmLx%, ToLx%

Friend Property Get Init(FmLx%, ToLx%) As FmToLx
If FmLx < 0 Then Stop
If ToLx < 0 Then Stop
If FmLx > ToLx - 1 Then Stop
Me.FmLx = FmLx
Me.ToLx = ToLx
Set Init = Me
End Property

Property Get ToStr$()
ToStr = FmtQQ("FmToLx(Fm(?) To(?))", FmLx, ToLx)
End Property

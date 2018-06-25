VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FmToPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public FmPos&, ToPos&

Property Get ToStr$()
ToStr = FmtQQ("(FmToPos ? ?)", FmPos, ToPos)
End Property

Friend Property Get Init(FmPos, ToPos) As FmToPos
If FmPos >= 0 And ToPos >= 0 And ToPos >= FmPos Then
    Me.FmPos = FmPos
    Me.ToPos = ToPos
Else
    Me.FmPos = 0
    Me.ToPos = 0
End If
Set Init = Me
End Property

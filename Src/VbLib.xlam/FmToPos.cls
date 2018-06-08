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

Sub Dmp()
Debug.Print ToStr
End Sub
Friend Property Get Init(FmPos&, ToPos&) As FmToPos
Me.FmPos = FmPos
Me.ToPos = ToPos
Set Init = Me
End Property


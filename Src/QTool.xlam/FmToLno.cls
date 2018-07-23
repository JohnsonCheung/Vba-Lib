VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FmToLno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public FmLno%, ToLno%

Friend Property Get Init(FmLno%, ToLno%) As FmToLno
If FmLno <= 0 Then Stop
If ToLno <= 0 Then Stop
If FmLno - 1 > ToLno Then Stop
Me.FmLno = FmLno
Me.ToLno = ToLno
Set Init = Me
End Property

Property Get ToStr$()
ToStr = FmtQQ("FmToLno(Fm(?) To(?))", FmLno, ToLno)
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FmCnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public FmLno%, Cnt%
Friend Property Get Init(FmLno, Cnt) As FmCnt
If FmLno <= 0 Then Stop
If Cnt <= 0 Then Stop
Me.FmLno = FmLno
Me.Cnt = Cnt
Set Init = Me
End Property
Property Get Zero() As FmCnt
Static O As New FmCnt, X As Boolean
If Not X Then
    X = True
    O.FmLno = -1
    O.Cnt = -1
End If
Set Zero = O
End Property
Property Get ToStr$()
ToStr = "FmCnt(FmLno(" & FmLno & ") Cnt(" & Cnt & "))"
End Property


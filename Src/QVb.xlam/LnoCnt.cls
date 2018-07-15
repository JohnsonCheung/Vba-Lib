VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LnoCnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Lno&, Cnt&

Property Get Init(Lno&, Cnt&) As LnoCnt
Me.Lno = Lno
Me.Cnt = Cnt
Set Init = Me
End Property

Property Get ToStr$()
ToStr = Tag("LnoCnt", Lno & " " & Cnt)
End Property

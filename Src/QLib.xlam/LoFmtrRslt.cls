VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoFmtrRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private B_Ly$(), B_Er As New Er, B_FmtWs As New FmtWs
Friend Property Get Init(Ly$(), Er As Er, FmtWs As FmtWs) As LoFmtrRslt
B_Ly = Ly
Set B_Er = Er
Set B_FmtWs = FmtWs
Set Init = Me
End Property

Property Get Ly() As String()
Ly = B_Ly
End Property

Property Get FmtWs() As FmtWs
Set FmtWs = B_FmtWs
End Property

Property Get Er() As Er
Set Er = B_Er
End Property

Property Get ToStr$()
ToStr = Tag_NyStr_ObjAp("LoFmtrRslt", "Ly FmtWs Er", SyObj(B_Ly), B_FmtWs, B_Er)
End Property

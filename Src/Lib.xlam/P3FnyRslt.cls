VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "P3FnyRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_Fny$()
Private A_Er$()
Private A_Ok$()
Function Init(Fny$(), Ok$(), Er$()) As P3FnyRslt
A_Fny = Fny
A_Er = Er
A_Ok = Ok
Set Init = Me
End Function
Property Get Show() As String()
Dim O$()
Push O, "P3FnyRslt("
Push O, FmtQQ("Fny(?)", JnSpc(A_Fny))
PushAy O, Vb.ErShow(A_Er)
Push O, "P3FnyRslt)"
Show = O
End Property
Property Get Fny() As String()
Fny = A_Fny
End Property
Property Get Er() As String()
Er = A_Er
End Property
Property Get Ok() As String()
Ok = A_Ok
End Property


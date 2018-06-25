VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Er"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private B_Er$()
Sub Push(A As Er)
PushAy B_Er, A.ErLy
End Sub

Property Get HasEr() As Boolean
HasEr = Sz(B_Er) > 0
End Property

Property Get ToStr$()
ToStr = Tag("Er", JnCrLf(B_Er, WithIx:=True))
End Property

Property Get ErLy() As String()
ErLy = B_Er
End Property

Sub PushMsg(Msg$)
VbAy.Push B_Er, Msg
End Sub

Sub PushErLy0(ErLy0)
PushAy B_Er, DftLy(ErLy0)
End Sub

Sub PushAp(ParamArray ErAp())
Dim I, Er As Er
For Each I In ErAp
    Set Er = I
    Push Er
Next
End Sub

Sub PushErLy0Ap(ParamArray ErLy0Ap())
Dim ErLy0
For Each ErLy0 In ErLy0Ap
    PushErLy0 ErLy0
Next
End Sub

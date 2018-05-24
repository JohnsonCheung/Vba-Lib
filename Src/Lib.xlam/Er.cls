VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Er"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private B_Er$()
Sub Add(A As Er)
PushAy B_Er, A.Val
End Sub
Property Get HasEr() As Boolean
HasEr = Sz(B_Er) > 0
End Property
Property Get ToStr$()
ToStr = JnCrLf(B_Er)
End Property
Property Get Val() As String()
Val = B_Er
End Property
Property Let Val(A$())
B_Er = A
End Property
Sub AddMsg(Msg$)
Push B_Er, Msg
End Sub

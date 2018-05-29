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
PushAy B_Er, A.ErLy
End Sub
Sub PushAp(ParamArray ErAp())
Stop
End Sub
Property Get HasEr() As Boolean
HasEr = Sz(B_Er) > 0
End Property
Property Get ToStr$()
If HasEr Then
    ToStr = "Er(" & vbCrLf & JnCrLf(B_Er) & vbCrLf & "Er)"
Else
    ToStr = "Er()"
End If
End Property
Property Get ErLy() As String()
ErLy = B_Er
End Property
Sub AddMsg(Msg$)
Push B_Er, Msg
End Sub
Sub AddErLy(A$())
PushAy B_Er, A
End Sub
Function AddErLyAp(ParamArray ErLyAp()) As Er
Dim Av(): Av = ErLyAp
Dim I, ErLy$()
For Each I In ErLyAp
    ErLy = I
    AddErLy ErLy
Next
Set AddErLyAp = Me
End Function

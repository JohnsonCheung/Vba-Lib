VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CnoInts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private B_Ay() As CnoInt
Sub Add(Cno%, I%)
Dim M As New CnoInt
M.Cno = Cno
M.I = I
PushObj B_Ay, M
End Sub
Property Get CnoIntAy() As CnoInt()
CnoIntAy = B_Ay
End Property
Property Get ToStr$()

Dim A$, O$()
For J = 0 To UB(B_Ay)
    Push O, B_Ay(J).ToStr
Next
ToStr = "CnoInts(" & vbCrLf & A & vbCrLf & ")"
End Property

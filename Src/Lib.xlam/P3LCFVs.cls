VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "P3LCFVs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private A_Ay() As P3LCFV
Private A_L_Lx%()
Private A_L_IsVF() As Boolean
Function Init(A() As P3LCFV, L_IsVF() As Boolean)
A_Ay = A
A_L_Lx = Oy.SrtedUniqIntAy("Lx")
If Sz(A_L_Lx) <> Sz(L_IsVF) Then PrmEr

Set Init = Me
End Function
Private Function Oy() As Oy
Dim O  As New Oy
Set Oy = O.Init(A_Ay)
End Function
Sub Push(A As P3LCFV)
PushObj A_Ay, A
End Sub
Property Get IsEmp() As Boolean
IsEmp = Sz(A_Ay) = 0
End Property
Function Ay() As P3LCFV()
Ay = A_Ay
End Function
Property Get Show() As String()
If IsEmp Then Exit Property
Dim O$()
VbAy.Push O, "P3LCFVs(----------------"
For Each I In A_Ay
    PushAy O, Cv(I).Show
Next
VbAy.Push O, "P3LCFVs)----------------"
End Property
Private Function Cv(I) As P3LCFV
Set Cv = I
End Function

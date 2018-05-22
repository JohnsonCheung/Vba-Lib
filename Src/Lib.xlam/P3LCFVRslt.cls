VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "P3LCFVRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private A_LCFVs As New P3LCFVs
Private A_Er() As String
Function Init(Ay() As P3LCFV, Er$()) As P3LCFVRslt
If Sz(Ay) > 0 Then
    If TypeName(Ay(0)) <> "P3LCFV" Then Stop
End If
A_LCFVAy = Ay
A_Er = Er
Set Init = Me
End Function
Property Get Er() As String()
Er = A_Er
End Property
Property Get LCFVAy() As P3LCFV()
LCFVAy = A_LCFVs.Ay
End Property
Property Get Drs() As Drs
With Drs
    .Fny = Fny
    .Dry = Dry
End With
End Property
Function FldValLy(T1$) As String()
Dim O$()
For J = 0 To UB(A_LCFVAy)
    With A_LCFVAy(J)
        Push O, T1 & " " & .Fld & " " & .Val
    End With
Next
FldValLy = O
End Function

Function ValFldLvsLy(T1$) As String()
Dim O$(), ValAy$(), V$, J%, I%, Fny$(), L%
ValAy = AyUniq(Oy.PrpSy("Val"))
For J = 0 To UB(ValAy)
    V = ValAy(J)
    For I = 0 To UB(A_LCFVAy)
        With A_LCFVAy(J)
            If .Val = V Then
                Push Fny, .Fld
            End If
        End With
    Next
    Push O, T1 & " " & V & " " & JnSpc(Fny)
Next
ValFldLvsLy = O
End Function
Function Show() As String()
Dim O$()
     Push O, "P3LCFVRslt(-----------------"
   PushAy O, A_LCFVs.Show
   PushAy O, ErShow(A_Er)
     Push O, "P3LCFVRslt)-----------------"
Show = O
End Function
Friend Sub Brw()
AyBrw Show
End Sub
Private Function Fny() As String()
Static Y$()
If Sz(Y) = 0 Then Y = LvsSy("Lx Cno Fld Val")
Fny = Y
End Function
Function Dry() As Variant()
Dry = Oy.Dry
End Function
Private Function Oy() As Oy
Set Oy = A_Vb.Oy(A_LCFVAy)
End Function
Friend Sub Tst()
Init ZZAy, ZZEr
AyBrw ValFldLvsLy("T1")
AyBrw FldValLy("T2")
End Sub
Private Function ZZAy() As P3LCFV()
Dim O() As P3LCFV
PushObj O, P3LCFV(1, 2, "sdf", "sdlfk")
PushObj O, P3LCFV(1, 2, "sdf", "sdlfk")
PushObj O, P3LCFV(1, 2, "sdf", "sdlfk")
ZZAy = O
End Function
Private Function ZZEr() As String()
ZZEr = ApSy("Lx(3) Is error")
End Function


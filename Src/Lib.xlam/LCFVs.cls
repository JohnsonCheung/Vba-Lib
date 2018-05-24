VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LCFVs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private B_Items As New VBA.Collection

Function Add(A As LCFVs) As LCFVs
Dim J%
'For J = 0 To A.U
'Next
End Function

Sub AddItem(A As LCFV)
B_Items.Add A
End Sub

Sub AddLCFV(Lx%, Cno%, F$, V)
Dim O As New LCFV
With O
    .Lx = Lx
    .Cno = Cno
    .F = F
    .V = V
End With
AddItem O
End Sub

Property Get Ay() As LCFV
Ay = Coll(B_Items).Ay
End Property

Property Get Drs() As Drs
With Drs
    .Fny = LvsSy("Lx Cno Fld Val")
    .Dry = Oy.Dry
End With
End Property

Private Property Get Oy() As Oy
Dim O As New Oy
Set Oy = O.InitByColl(B_Items)
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CnoVals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private B_Ay() As CnoVal

Property Get Ay() As CnoVal()
Ay = B_Ay
End Property

Property Get CnoAy() As Integer()
CnoAy = Oy.PrpIntAy("Cno")
End Property

Private Function Oy() As Oy
Set Oy = Vb.Oy(B_Ay)
End Function
Property Get ValAy() As Variant()
ValAy = Oy.PrpAy("V")
End Property
Function StrValAy() As String()
StrValAy = Oy.PrpSy("V")
End Function
Function CnoIx%(Cno%)
'Use Cno to find any element in B_Ay has .Cno = Cno,
'Return the Ix of B_Ay if found else return -1
Dim J%
For J = 0 To U
    If B_Ay(J).Cno = Cno Then CnoIx = J: Exit Function
Next
CnoIx = -1
End Function

Sub AddCnoVal(Cno%, F$, V)
Dim O As New CnoVal
With O
    .F = F
    .Cno = Cno
    .V = V
End With
AddItm O
End Sub

Sub AddItm(A As CnoVal)
PushObj B_Ay, A
End Sub

Property Get N&()
N = Sz(B_Ay)
End Property

Property Get U&()
U = N - 1
End Property

Property Get IsEmp() As Boolean
IsEmp = N = 0
End Property

Property Get ToStr$()
ToStr = Tag("CnoVals", Oy.ToStr)
End Property


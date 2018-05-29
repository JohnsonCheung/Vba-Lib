VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CnosLvls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private B_Ay() As CnosLvl
Sub Fmt(A As Range)
Dim I, M As CnosLvl
For Each I In B_Ay
    Set M = I
    M.Fmt A
Next
End Sub
Property Get CnosLvlAy() As CnosLvl()
CnosLvlAy = B_Ay
End Property
Property Let CnosLvlAy(V() As CnosLvl)
B_Ay = V
End Property

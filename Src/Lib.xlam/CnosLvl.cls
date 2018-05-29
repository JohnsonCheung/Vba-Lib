VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CnosLvl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private B_CnoAy%()
Public Lvl As Byte

Sub SetCnoAyLvl(CnoAy%(), L%)
B_CnoAy = CnoAy
Lvl = L%
End Sub

Sub Fmt(A As Range)
Dim Cno
If IsEmp Then Exit Sub
For Each Cno In B_CnoAy
    RgC(A, Cno).EntireColumn.OutlineLevel = Lvl
Next
End Sub

Property Get IsEmp() As Boolean
IsEmp = N = 0
End Property

Property Get N%()
N = Sz(B_CnoAy)
End Property

Property Get CnoAy() As Integer()
CnoAy = B_CnoAy
End Property

Property Let CnoAy(V%())
B_CnoAy = V
End Property


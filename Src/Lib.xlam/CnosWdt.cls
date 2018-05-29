VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CnosWdt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private B_CnoAy%()
Public Wdt%

Sub SetCnoAyWdt(CnoAy%(), W%)
B_CnoAy = CnoAy
Wdt = W%
End Sub

Sub Fmt(A As Range)
Dim Cno
If IsEmp Then Exit Sub
For Each Cno In B_CnoAy
    RgC(A, Cno).Width = Wdt
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

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CnoAyWdtAy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private B_CnoAy%(), B_WdtAy%()
Property Get CnoAy() As Integer()
CnoAy = B_CnoAy
End Property
Property Get WdtAy() As Integer()
WdtAy = B_WdtAy
End Property
Sub SetCnoAyWdtAy(CnoAy%(), Wdt%())
If Sz(CnoAy) <> Sz(Wdt) Then PmEr
B_CnoAy = CnoAy
B_WdtAy = WdtAy
End Sub

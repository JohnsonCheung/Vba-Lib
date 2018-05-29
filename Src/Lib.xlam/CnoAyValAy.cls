VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CnoAyValAy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private B_CnoAy%()
Private B_ValAy()
Sub SetCnoAyValAy(CnoAy%(), ValAy())
B_CnoAy = CnoAy
B_ValAy = ValAy
End Sub

Property Get CnoAy() As Integer()
CnoAy = B_CnoAy
End Property
Property Get ValAy() As Variant()
ValAy = B_ValAy
End Property

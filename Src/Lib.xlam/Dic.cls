VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private B_Dic As Dictionary
Friend Property Get Init(A As Dictionary) As Dic
Set B_Dic = A
Set Init = Me
End Property

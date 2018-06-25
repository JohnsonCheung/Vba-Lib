VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Drs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private B_Fny$()
Private B_Dry()
Property Get Init(Fny0, Dry()) As Drs
B_Dry = Dry
B_Fny = DftNy(Fny0)
Set Init = Me
End Property
Property Get Dry() As Variant()
Dry = B_Dry
End Property
Property Get Fny() As String()
Fny = B_Fny
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public DtNm$
Private A_Fny$()
Private A_Dry() As Variant

Property Get Dry() As Variant()
Dry = A_Dry
End Property

Property Get Fny() As String()
Fny = A_Fny
End Property

Friend Property Get Init(DtNm$, Fny0, Dry()) As Dt
A_Dry = Dry
A_Fny = DftNy(Fny0)
A_DtNm = DtNm
Set Init = Me
End Property

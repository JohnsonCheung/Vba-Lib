VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Drs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private A_Dry() As Variant
Private A_Fny() As String

Property Get Dry() As Variant()
Dry = A_Dry
End Property

Property Get Fny() As String()
Fny = A_Fny
End Property

Friend Property Get Init(Fny0, Dry()) As Drs
A_Fny = DftNy(Fny0)
A_Dry = Dry
Set Init = Me
End Property

Property Get ToStr$()
Dim A$, B$
'A = Join(A_Fny, " ")
'ToStr = "Drs(|Fny(" & A & ")|" & B & "|)"
', Join(A_Fny, " "), DryToStr(A_Dry))
End Property

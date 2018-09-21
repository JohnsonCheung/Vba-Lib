VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FmTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Fmix&, Toix&

Friend Property Get Init(Fmix&, Toix&) As FmTo
Me.Fmix = Fmix
Me.Toix = Toix
Set Init = Me
End Property

Property Get ToStr$()
ToStr = FmtQQ("FmTo(? ?)", Fmix, Toix)
End Property

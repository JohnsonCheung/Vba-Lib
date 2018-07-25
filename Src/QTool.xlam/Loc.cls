VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Loc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public PjNm$, MdNm$, LCC As LCC
Friend Property Get Init(PjNm$, MdNm$, LCC As LCC) As Loc
Me.PjNm = PjNm

End Property

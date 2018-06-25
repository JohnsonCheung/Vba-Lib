VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CnoVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Cno%, F$, V

Property Get ToStr$()
ToStr = FmtQQ("CnoVal(? V-Ty(?) Fld(?) V(?))", Cno, TypeName(V), F, V)
End Property

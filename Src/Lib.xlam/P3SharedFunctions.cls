VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "P3SharedFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Function LxLin_P3Lin(Lx%, Lin$, Optional IsVF As Boolean) As P3Lin
Dim A$, B$, C$
LinAsgTTRst Lin, A, B, C
Set LxLin_P3Lin = P3Lin(Lx, B, C)
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "P3Lin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Lx%, B$, C$, Par As P3Itm
Property Get IsVF() As Boolean
IsVF = Par.IsVF
End Property
Property Get Fny() As String()
If IsVF Then
    Fny = LvsSy(C)
Else
    Fny = ApSy(B)
End If
End Property
Function Show() As String()
Show = ApSy(FmtQQ("Lx B C = [?] [?] [?]", Lx, B, C))
End Function
Property Get Lin$()
With Par
    Lin = Lx & " " & .Nm & " " & B & " " & C
End With
End Property


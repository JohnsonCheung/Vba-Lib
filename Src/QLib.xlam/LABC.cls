VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LABC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Par As LABCs, Lx%, B$, C$

Property Get A$()
If IsNothing(Par) Then A = "?": Exit Property
Dim T1$: T1 = Par.T1
If T1 = "" Then
    A = "?"
Else
    A = T1
End If
End Property

Property Get FldLvs() As String
FldLvs = IIf(IsVF, C, B)
End Property

Property Get Fny() As String()
Fny = LvsSy(FldLvs)
End Property

Property Get IsVF() As Boolean
IsVF = Par.IsVF
End Property

Property Get Lin$()
Lin = A & " " & B & " " & C
End Property

Property Get ToStr$()
ToStr = FmtQQ("IsVF(?) LABC(? ? ? ?)", Par.IsVF, Lx, A, B, C)
End Property

Property Get Val$()
If IsVF Then Val = B Else Val = C
End Property

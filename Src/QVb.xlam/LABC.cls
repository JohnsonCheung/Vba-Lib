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
Public Lx%, A$, B$, C$, IsVF As Boolean

Property Get ToStr$()
ToStr = FmtQQ("LABC(IsVF(?) Lx(?) A(?) B(?) C(?))", IsVF, Lx, A, B, C)
End Property

Friend Property Get Init(Lx%, A$, B$, C$, IsVF As Boolean) As LABC
With Me
    .Lx = Lx
    .A = A
    .B = B
    .C = C
    .IsVF = IsVF
End With
Set Init = Me
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lnx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private B_Lx%, B_Lin$
Property Get Lx%()
Lx = B_Lx
End Property
Property Get Lin$()
Lin = B_Lin
End Property
Function Fmt$(QQVblStr$)
Dim A$: A = RplVBar(QQVblStr)
Fmt = FmtQQ(A, B_Lin, B_Lx)
End Function
Friend Property Get Init(Lin$, Lx%) As Lnx
B_Lx = Lx
B_Lin = Lin
Set Init = Me
End Property

Property Get ToStr$()
ToStr = FmtQQ("Lnx(?: ?", B_Lx, B_Lin)
End Property

Property Get RmvDDRmk() As Lnx
Dim O As New Lnx
Dim L$:
    L = Vb.Lin(B_Lin).RmvDDRmk.Lin
Set RmvDDRmk = O.Init(L, B_Lx)
End Property

Property Get IsBlankLin() As Boolean
IsBlankLin = Trim(B_Lin) = ""
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StmtsRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private A_Stmts$
Private A_ErLy$()

Property Get ErLy() As String()
ErLy = B_ErLy
End Property

Property Get Stmts$()
Stmts = A_Stmts
End Property

Property Get ToStr$()
Stop '
End Property

Friend Property Get Init(Stmts$, ErLy$()) As StmtsRslt
A_Stmts = Stmts
A_ErLy = ErLy
End Property

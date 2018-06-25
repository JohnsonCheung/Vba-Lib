VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StmtRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Stmt$
Private B_ErLy$()
Friend Property Get Init(Stmt$, ErLy$()) As StmtRslt
Me.Stmt = Stmt
B_ErLy = ErLy
End Property

Property Get ErLy() As String()
ErLy = B_ErLy
End Property

Property Get ToStr$()
Stop '
End Property

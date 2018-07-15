VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StmtRslt1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Stmt$
Private B_ErLy$()

Property Get ErLy() As String()
ErLy = B_ErLy
End Property

Property Let ErLy(V$())
B_ErLy = V
End Property

Friend Property Get Init(A() As Lnx1) As StmtRslt

End Property

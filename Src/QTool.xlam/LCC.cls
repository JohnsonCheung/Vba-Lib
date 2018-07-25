VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Lno%, C1%, C2%
Friend Property Get Init(Lno%, C1%, C2%) As LCC
Me.Lno = Lno
Me.C1 = C1
Me.C2 = C2
Set Init = Me
End Property

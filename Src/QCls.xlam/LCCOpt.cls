VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LCCOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Som As Boolean, LCC As LCC
Friend Property Get Init(Som As Boolean, LCC As LCC) As LCCOpt
If Som Then
    Set Me.LCC = LCC
    If ObjPtr(LCC) = 0 Then Stop
End If
Me.Som = Som
Set Init = Me
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LCFVsRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public LCFVs As LCFVs
Private B_Er$()

Property Let Er(A$())
B_Er = A
End Property

Property Get Er() As String()
Er = B_Er
End Property

Sub AddMsg(Msg$)
Push B_Er, Msg
End Sub

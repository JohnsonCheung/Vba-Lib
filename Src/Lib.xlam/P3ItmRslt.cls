VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "P3ItmRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_Itm As New P3Itm
Private A_Er() As String
Function Init(A As P3Itm, Er$()) As P3ItmRslt
Set A_Itm = A
A_Er = Er
Set Init = Me
End Function
Property Get Er() As String()
Er = A_Er
End Property
Property Get Lins() As P3Itm
'Set Lins = P3Itm
End Property
Function Show() As String()
Dim O$()
     Push O, "P3ItmRslt("
   PushAy O, A_Itm.Show
   PushAy O, ErShow(A_Er)
     Push O, "P3ItmRslt)"
Show = O
End Function
Friend Sub Brw()
AyBrw Show
End Sub

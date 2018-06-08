VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Collx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A As VBA.Collection
Property Get Init(Coll As VBA.Collection) As Collx
Set A = Coll
Set Init = Me
End Property
Function AyAs(OAsAy)
Dim O: O = OAsAy: Erase O
Dim I
For Each I In A
    Push O, I
Next
AyAs = O
End Function
Function Ay() As Variant()
Dim O(), I
For Each I In A
    Push O, I
Next
Ay = O
End Function

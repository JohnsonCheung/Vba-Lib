VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Coll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private B_Coll As VBA.Collection
Function Init(A As VBA.Collection) As Coll
Set B_Coll = A
Set Init = Me
End Function
Function AyAs(OAsAy)
Dim O: O = OAsAy: Erase O
Dim I
For Each I In B_Coll
    Push O, I
Next
AyAs = O
End Function
Function Ay() As Variant()
Dim O(), I
For Each I In B_Coll
    Push O, I
Next
Ay = O
End Function

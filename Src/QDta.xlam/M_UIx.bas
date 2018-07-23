Attribute VB_Name = "M_UIx"
Option Explicit

Function UIxAy(U&) As Long()
Dim O&(), J&
ReDim O(U)
For J = 0 To U
    O(J) = J
Next
UIxAy = O
End Function

Attribute VB_Name = "M_Itr"
Option Explicit

Property Get ItrAy(A) As Variant()
Dim O(), I
For Each I In A
    Push O, I
Next
ItrAy = O
End Property

Property Get ItrCast(A, CastToAy)
Dim O: O = CastToAy: Erase O
Dim I
For Each I In A
    Push O, I
Next
ItrCast = O
End Property

Property Get ItrNy(A) As String()
Dim O$(), I
For Each I In A
    Push O, ObjNm(I)
Next
ItrNy = O
End Property

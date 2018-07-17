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

Property Get ItrPrpValAy(A, PrpNm) As Variant()
Dim O(), I
For Each I In A
    Push O, CallByName(I, PrpNm, VbGet)
Next
ItrPrpValAy = O
End Property
Property Get ItrPrpSy(A, PrpNm) As String()
Dim O$(), I
For Each I In A
    Push O, CallByName(I, PrpNm, VbGet)
Next
ItrPrpSy = O
End Property
Property Get ItrFstNm$(A)
Dim I
For Each I In A
    ItrFstNm = ObjNm(I)
Next
End Property
Property Get ItrNy(A, Optional Patn$ = ".") As String()
Dim O$(), I
If Patn = "." Then
    For Each I In A
        Push O, ObjNm(I)
    Next
Else
    Dim R As RegExp: Set R = Re(Patn)
    Dim N$
    For Each I In A
        N = ObjNm(I)
        If R.Test(N) Then
            Push O, N
        End If
    Next
End If
ItrNy = O
End Property
Property Get ItrHasNm(A, Nm) As Boolean
Dim I
For Each I In A
    If I.Name = Nm Then ItrHasNm = True: Exit Property
Next
End Property


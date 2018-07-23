Attribute VB_Name = "M_Itr"
Option Explicit

Function ItrAy(A) As Variant()
Dim O(), I
For Each I In A
    Push O, I
Next
ItrAy = O
End Function

Function ItrCast(A, CastToAy)
Dim O: O = CastToAy: Erase O
Dim I
For Each I In A
    Push O, I
Next
ItrCast = O
End Function

Function ItrCntByBoolPrp&(A, BoolPrpNm$)
If A.Count = 0 Then Exit Function
Dim O, Cnt&
For Each O In A
    If CallByName(O, BoolPrpNm, VbGet) Then
        Cnt = Cnt + 1
    End If
Next
ItrCntByBoolPrp = Cnt
End Function

Function ItrFstNm$(A)
Dim I
For Each I In A
    ItrFstNm = ObjNm(I)
Next
End Function

Function ItrHasNm(A, Nm) As Boolean
Dim I
For Each I In A
    If I.Name = Nm Then ItrHasNm = True: Exit Function
Next
End Function

Function ItrNy(A, Optional Patn$ = ".") As String()
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
End Function

Function ItrPrpSy(A, PrpNm) As String()
Dim O$(), I
For Each I In A
    Push O, CallByName(I, PrpNm, VbGet)
Next
ItrPrpSy = O
End Function

Function ItrPrpValAy(A, PrpNm) As Variant()
Dim O(), I
For Each I In A
    Push O, CallByName(I, PrpNm, VbGet)
Next
ItrPrpValAy = O
End Function

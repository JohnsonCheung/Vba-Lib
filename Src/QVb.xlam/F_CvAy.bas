Attribute VB_Name = "F_CvAy"
Option Explicit

Function AyBoolAy(A) As Boolean()
AyBoolAy = AyInto(A, EmpBoolAy)
End Function

Function AyBytAy(A) As Byte()
AyBytAy = AyInto(A, EmpBytAy)
End Function

Function AyInto(A, OInto)
Ass IsArray(A)
Ass IsArray(OInto)
If TypeName(A) = TypeName(OInto) Then
    AyInto = A
    Exit Function
End If
Dim O: O = OInto: Erase O
Dim I
If Sz(A) = 0 Then AyInto = O: Exit Function
For Each I In A
    Push O, I
Next
AyInto = O
End Function

Function AyDblAy(A) As Double()
AyDblAy = AyInto(A, EmpDblAy)
End Function

Function AyIntAy(A) As Integer()
AyIntAy = AyInto(A, EmpIntAy)
End Function

Function AyLngAy(A) As Long()
AyLngAy = AyInto(A, EmpLngAy)
End Function

Function AySngAy(A) As Single()
AySngAy = AyInto(A, EmpSngAy)
End Function

Function AySy(A) As String()
AySy = AyInto(A, EmpSy)
End Function

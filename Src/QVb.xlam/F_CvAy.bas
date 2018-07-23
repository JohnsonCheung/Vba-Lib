Attribute VB_Name = "F_CvAy"
Option Explicit

Function AyBoolAy(Ay) As Boolean()
AyBoolAy = AyCast(Ay, EmpBoolAy)
End Function

Function AyBytAy(Ay) As Byte()
AyBytAy = AyCast(Ay, EmpBytAy)
End Function

Function AyCast(Ay, IntoAy)
Ass IsArray(Ay)
Ass IsArray(IntoAy)
If TypeName(Ay) = TypeName(IntoAy) Then
    AyCast = Ay
    Exit Function
End If
Dim O: O = IntoAy: Erase O
Dim I
For Each I In Ay
    Push O, I
Next
AyCast = O
End Function

Function AyDblAy(Ay) As Double()
AyDblAy = AyCast(Ay, EmpDblAy)
End Function

Function AyIntAy(Ay) As Integer()
AyIntAy = AyCast(Ay, EmpIntAy)
End Function

Function AyLngAy(Ay) As Long()
AyLngAy = AyCast(Ay, EmpLngAy)
End Function

Function AySngAy(Ay) As Single()
AySngAy = AyCast(Ay, EmpSngAy)
End Function

Function AySy(Ay) As String()
AySy = AyCast(Ay, EmpSy)
End Function

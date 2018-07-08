Attribute VB_Name = "F_CvAy"
Option Explicit

Property Get AyBoolAy(Ay) As Boolean()
AyBoolAy = AyCast(Ay, EmpBoolAy)
End Property

Property Get AyBytAy(Ay) As Byte()
AyBytAy = AyCast(Ay, EmpBytAy)
End Property

Property Get AyCast(Ay, IntoAy)
Ass IsArray(Ay)
Ass IsArray(IntoAy)
If TypeName(Ay) = TypeName(IntoAy) Then
    AyCast = Ay
    Exit Property
End If
Dim O: O = IntoAy: Erase O
Dim I
For Each I In Ay
    Push O, I
Next
AyCast = O
End Property

Property Get AyIntAy(Ay) As Integer()
AyIntAy = AyCast(Ay, EmpIntAy)
End Property

Property Get AyLngAy(Ay) As Long()
AyLngAy = AyCast(Ay, EmpLngAy)
End Property

Property Get AySngAy(Ay) As Single()
AySngAy = AyCast(Ay, EmpSngAy)
End Property

Property Get AySy(Ay) As String()
AySy = AyCast(Ay, EmpSy)
End Property

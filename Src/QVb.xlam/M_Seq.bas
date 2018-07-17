Attribute VB_Name = "M_Seq"
Option Explicit
Property Get SeqOf__(FmNum, ToNum, OAy)
Dim O&()
ReDim OAy(Abs(FmNum - ToNum))
Dim J&, I&
If ToNum > FmNum Then
    For J = FmNum To ToNum
        OAy(I) = J
        I = I + 1
    Next
Else
    For J = ToNum To FmNum Step -1
        OAy(I) = J
        I = I + 1
    Next
End If
SeqOf__ = OAy
End Property
Property Get SeqOfInt(FmNum%, ToNum%) As Integer()
SeqOfInt = SeqOf__(FmNum, ToNum, EmpIntAy)
End Property
Property Get SeqOfLng(FmNum&, ToNum&) As Long()
SeqOfLng = SeqOf__(FmNum, ToNum, EmpLngAy)
End Property

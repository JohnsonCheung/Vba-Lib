Attribute VB_Name = "M_Fml"
Option Explicit
Property Get FmlErMsgOpt(Fml$, Fny$()) As SomStr
Dim A$(): A = Ny
Dim O$(), J%
For J = 0 To UB(A)
    If Not AyHas(Fny, A(J)) Then
        Push O, A(J)
    End If
Next
If AyIsEmp(O) Then Exit Property
FmlErMsgOpt = SomStr(JnSpc(O))
End Property

Property Get FmlNy(Fml$) As String()
FmlNy = Macro(A).Ny(ExclBkt:=True, Bkt:="[]")
End Property



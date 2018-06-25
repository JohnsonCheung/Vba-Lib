VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Fml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A$
Friend Property Get Init(Fml$) As Fml
A = Fml
Set Init = Me
End Property
Function ErMsgOpt(Fny$()) As StrOpt
Dim A$(): A = Ny
Dim O$(), J%
For J = 0 To UB(A)
    If Not AyHas(Fny, A(J)) Then
        Push O, A(J)
    End If
Next
If AyIsEmp(O) Then Exit Function
ErMsgOpt = SomStr(JnSpc(O))
End Function

Property Get Ny() As String()
Ny = Macro(A).Ny(ExclBkt:=True, Bkt:="[]")
End Property


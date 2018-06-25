Attribute VB_Name = "M_LnxAy"
Option Explicit

Property Get FmtLy(QQVblStr$) As String()
Dim O$(), J%
Stop '
'For J = 0 To U
'    M_Ay.Push O, B_Ay(J).Fmt(QQVblStr)
'Next
FmtLy = O
End Property

Property Get LnxAy_Ly(A() As Lnx) As String()
Dim O$(), J%
Stop '
'For J = 0 To U
'    M_Ay.Push O, B_Ay(J).Lin
'Next
LnxAy_Ly = O
End Property

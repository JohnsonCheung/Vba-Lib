Attribute VB_Name = "A_Vb"
Option Explicit
Public Const M_Fld_IsInValid$ = "Lx(?) Fld(?) is invalid.  Not found in Fny"
Public Const M_Fld_IsDup$ = "Lx(?) Fld(?) is found dup in Lx(?)."
Function IsStr(V) As Boolean
IsStr = VarType(V) = vbString
End Function
Property Get Cmd() As Cmd
Static Y As New Cmd
Set Cmd = Y
End Property
Property Get Oy(A) As Oy
Dim O As New Oy
Set Oy = O.Init(A)
End Property
Property Get V(A) As V
Dim O As New V
O.Init A
Set V = O
End Property
Function SrcLin(A) As SrcLin
Dim O As New SrcLin
O.Init A
Set SrcLin = O
End Function
Function Sy(A$()) As Sy
Dim O As New Sy
Set Sy = O.Init(A)
End Function
Function ErShow(Er$()) As String()
ErShow = SyShow("Er", Er)
End Function
Function OkShow(Ok$()) As String()
OkShow = SyShow("Ok", Ok)
End Function
Function SyShow(XX$, Sy$()) As String()
Dim O$()
Push O, XX & "(----------------------"
PushAy O, Sy
Push O, XX & ")----------------------"
SyShow = O
End Function
Sub PrmEr()
MsgBox "Prm Er"
Stop
End Sub

Attribute VB_Name = "Vb"
Option Explicit
Property Get Lin(A) As Lin
Dim O As New Lin
Set Lin = O.Init(A)
End Property
Property Get Fso() As Scripting.FileSystemObject
Static Y As New Scripting.FileSystemObject
Set Fso = Y
End Property
Property Get Lnx(Lin$, Lx%) As Lnx
Dim O As New Lnx
Set Lnx = O.Init(Lin, Lx)
End Property
Property Get Lines(A) As Lines
Dim O As New Lines
Set Lines = O.Init(A)
End Property
Property Get Ly(Ly0) As Ly
Dim O As New Ly
Dim L$(): L = Ly0
Set Ly = O.Init(L)
End Property
Property Get Pth(A) As Pth
Dim O As New Pth
Set Pth = O.Init(A)
End Property
Property Get Ffn(A) As Ffn
Dim O As New Ffn
Set Ffn = O.Init(A)
End Property
Property Get Emp() As Emp
Static Y As New Emp
Set Emp = Y
End Property
Property Get S1S2(S1, S2) As S1S2
Dim O As New S1S2
Set S1S2 = O.Init(S1, S2)
End Property
Property Get Macro(A) As Macro
Dim O As New Macro
Set Macro = O.Init(A)
End Property
Property Get FmTo(FmIx&, ToIx&) As FmTo
Dim O As New FmTo
Set FmTo = O.Init(FmIx, ToIx)
End Property
Property Get Ft(A) As Ft
Dim O As New Ft
Set Ft = O.Init(A)
End Property
Property Get Re(RegExpStr$) As Re
Dim O As New Re
Set Re = O.Init(RegExpStr)
End Property

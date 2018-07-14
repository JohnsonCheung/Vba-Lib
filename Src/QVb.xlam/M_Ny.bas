Attribute VB_Name = "M_Ny"
Option Explicit

Property Get NyIxLy(Ny0) As String()
'It is to return 2 lines with
'first line is 0   1     2 ..., where 0,1,2.. are ix of A$()
'second line is each element of A$() separated by A
'Eg, A$() = "A BBBB CCC DD"
'return 2 lines of
'0 1    2   3
'A BBBB CCC DD
Dim Ny$(): Ny = DftNy(Ny0)
If Sz(Ny) = 0 Then Exit Property
Dim A1$()
Dim A2$()
Dim U&: U = UB(Ny)
ReSz A1, U
ReSz A2, U
Dim O$(), J%, L$, W%
For J = 0 To U
    L = Len(Ny(J))
    W = Max(L, Len(J))
    A1(J) = AlignL(J, W)
    A2(J) = AlignL(Ny(J), W)
Next
Push O, JnSpc(A1)
Push O, JnSpc(A2)
NyIxLy = O
End Property

Attribute VB_Name = "IdeMthBrk"
Option Explicit
Sub LinMthBrkAsg(A, OShtMdy$, OShtTy$, ONm$)
Dim L$: L = A
OShtMdy = ShfShtMdy(L)
OShtTy = ShfMthShtTy(L)
ONm = TakNm(L)
End Sub

Function LinMthBrk(A) As String()
LinMthBrk = ShfMthBrk(CStr(A))
End Function


Function IsNmSel(Nm$, Re As RegExp, ExlLikAy) As Boolean
If Nm = "" Then Exit Function
If Not IsNothing(Re) Then
    If Not Re.Test(Nm) Then Exit Function
End If
IsNmSel = Not IsInLikAy(Nm, ExlLikAy)
End Function

Function ShfMthBrk(OLin$) As String()
ReDim B$(2)
B(0) = ShfShtMdy(OLin)
B(1) = ShfMthShtTy(OLin): If B(1) = "" Then ShfMthBrk = B: Exit Function
B(2) = ShfNm(OLin)
ShfMthBrk = B
End Function


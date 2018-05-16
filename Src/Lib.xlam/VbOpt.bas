Attribute VB_Name = "VbOpt"
Option Explicit
Type VarOpt
   Som As Boolean
   V As Variant
End Type
Type SyOpt
   Som As Boolean
   Sy() As String
End Type
Type VOpt
   Som As Boolean
   V As Variant
End Type
Type DicOpt
   Dic As Dictionary
   Som As Boolean
End Type

Sub X()
Dim I
For Each I In PjMdNyOfStd(CurPj)
    If Left(I, 2) = "D0" Then
        Md(I).Parent.Name = "Dta" & Mid(I, 4)
    End If
Next
End Sub

Function SomDic(A As Dictionary) As DicOpt
Set SomDic.Dic = A
SomDic.Som = True
End Function

Function SomStr(S) As StrOpt
SomStr.Som = True
SomStr.Str = S
End Function

Function SomSy(Sy$()) As SyOpt
SomSy.Som = True
SomSy.Sy = Sy
End Function

Function SomV(V) As VOpt
SomV.Som = True
SomV.V = V
End Function

Function SomVar(V) As VarOpt
SomVar.Som = True
SomVar.V = V
End Function

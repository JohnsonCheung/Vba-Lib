Attribute VB_Name = "VbStrRe"
Option Explicit

Function NewRe(Patn$, Optional MultiLine As Boolean, Optional IgnoreCase As Boolean, Optional IsGlobal As Boolean) As RegExp
Dim O As New RegExp
ReSet O, Patn, MultiLine, IgnoreCase, IsGlobal
Set NewRe = O
End Function

Sub ReSet(Re As RegExp, Patn$, Optional MultiLine As Boolean, Optional IgnoreCase As Boolean, Optional IsGlobal As Boolean)
With Re
   .Pattern = Patn
   .MultiLine = MultiLine
   .IgnoreCase = IgnoreCase
   .Global = IsGlobal
End With
End Sub

Function ReTst(S, Patn$, Optional MultiLine As Boolean, Optional IgnoreCase As Boolean, Optional IsGlobal As Boolean) As Boolean
Static Re As New RegExp
ReSet Re, Patn, MultiLine, IgnoreCase, IsGlobal
ReTst = Re.Test(S)
End Function

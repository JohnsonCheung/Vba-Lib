Attribute VB_Name = "M_OptAy"
Option Explicit

Function OptAy_HasNone(A) As Boolean
If Sz(A) = 0 Then Exit Function
OptAy_HasNone = True
Dim Opt
For Each Opt In A
    If Not Opt.Som Then Exit Function
Next
OptAy_HasNone = False
End Function

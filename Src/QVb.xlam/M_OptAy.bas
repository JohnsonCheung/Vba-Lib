Attribute VB_Name = "M_OptAy"
Option Explicit

Property Get OptAy_HasNone(A) As Boolean
If Sz(A) = 0 Then Exit Property
OptAy_HasNone = True
Dim Opt
For Each Opt In A
    If Not Opt.Som Then Exit Property
Next
OptAy_HasNone = False
End Property


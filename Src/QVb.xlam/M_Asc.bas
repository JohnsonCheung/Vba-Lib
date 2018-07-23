Attribute VB_Name = "M_Asc"
Option Explicit
Function AscIsDigit(A%) As Boolean
If A < 48 Then Exit Function
If A > 57 Then Exit Function
AscIsDigit = True
End Function

Function AscIsLCase(A%) As Boolean
If A < 97 Then Exit Function
If A > 122 Then Exit Function
AscIsLCase = True
End Function

Function AscIsUCase(A%) As Boolean
If A < 65 Then Exit Function
If A > 90 Then Exit Function
AscIsUCase = True
End Function


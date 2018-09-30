Attribute VB_Name = "VbFun"
Option Explicit

Function IsEqObj(A, B) As Boolean
IsEqObj = ObjPtr(A) = ObjPtr(B)
End Function


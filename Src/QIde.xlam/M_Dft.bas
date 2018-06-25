Attribute VB_Name = "M_Dft"
Option Explicit
Function DftMd(A As CodeModule) As CodeModule
If IsNothing(A) Then
   Set DftMd = CurMd
Else
   Set DftMd = A
End If
End Function
Function DftCmpTyAy(A) As vbext_ComponentType()
If IsLngAy(A) Then DftCmpTyAy = A
End Function

Function DftMdByMdNm(A$) As CodeModule
If A = "" Then
   Set DftMdByMdNm = CurMd
Else
   Set DftMdByMdNm = Md(DftMdNm(A))
End If
End Function

Function DftMdNm$(A$)
If A = "" Then
   DftMdNm = CurMdNm
Else
   DftMdNm = A
End If
End Function



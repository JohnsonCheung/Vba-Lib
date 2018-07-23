Attribute VB_Name = "M_Dft"
Option Explicit

Function DftCmpTyAy(A) As vbext_ComponentType()
If IsLngAy(A) Then DftCmpTyAy = A
End Function

Function DftMd(A As CodeModule) As CodeModule
If IsNothing(A) Then
   Set DftMd = CurMd
Else
   Set DftMd = A
End If
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

Function DftPj(A As VBProject) As VBProject
If IsNothing(A) Then
   Set DftPj = CurPj
Else
   Set DftPj = A
End If
End Function

Function DftPjByPjNm(A$) As VBProject
If A = "" Then
   Set DftPjByPjNm = CurPjx
   Exit Function
End If
Dim I As VBProject
For Each I In CurVbe.VBProjects
   If UCase(I.Name) = UCase(A) Then Set DftPjByPjNm = I: Exit Function
Next
Stop
End Function

Sub DftCmpTyAy__Tst()
Dim X() As vbext_ComponentType
DftCmpTyAy (X)
Stop
End Sub

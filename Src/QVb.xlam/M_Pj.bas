Attribute VB_Name = "M_Pj"
Option Explicit
Function Pj(A) As VBProject
Set Pj = CurVbe.VBProjects(A)
End Function

Function PjMd(A, MdNm) As CodeModule
Set PjMd = A.VBComponents(MdNm).CodeModule
End Function


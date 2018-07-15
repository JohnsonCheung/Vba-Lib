Attribute VB_Name = "M_Vbe"
Option Explicit

Property Get Vbe_PjAy(A As Vbe) As VBProject()
VbePjAy = ItrIntoAy(A.VBProjects, EmpPjAy)
End Property

Property Get Vbe_PjNy(A As Vbe) As String()
VbePjNy = OyNy(VbePjAy(A))
End Property

Sub Vbe_Export(A As Vbe)
Dim I
For Each I In Vbe_PjAy(A)
    Pj_Export CvPj(I)
Next
End Sub

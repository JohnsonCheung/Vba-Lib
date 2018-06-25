Attribute VB_Name = "M_Vbe"
Option Explicit
Sub VbeExport(A As Vbe)
Dim I
For Each I In VbePjAy(A)
    PjExport CvPj(I)
Next
End Sub
Property Get VbePjAy(A As Vbe) As VBProject()
VbePjAy = ItrIntoAy(A.VBProjects, EmpPjAy)
End Property
Property Get VbePjNy(A As Vbe) As String()
VbePjNy = OyNy(VbePjAy(A))
End Property


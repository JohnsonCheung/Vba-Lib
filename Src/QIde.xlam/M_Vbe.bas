Attribute VB_Name = "M_Vbe"
Option Explicit

Sub Vbe_Export(A As VBE)
Dim I
For Each I In Vbe_PjAy(A)
    Pj_Export CvPj(I)
Next
End Sub

Property Get Vbe_PjAy(A As VBE) As VBProject()
VbePjAy = ItrIntoAy(A.VBProjects, EmpPjAy)
End Property

Property Get Vbe_PjNy(A As VBE) As String()
VbePjNy = OyNy(VbePjAy(A))
End Property
Function VbeVisWinCnt%(A As VBE)
VbeVisWinCnt = ItrCntByBoolPrp(A.Windows, "Visible")
End Function

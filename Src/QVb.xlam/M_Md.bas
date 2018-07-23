Attribute VB_Name = "M_Md"
Option Explicit

Function Md(MdDNm) As CodeModule
Dim A1$(): A1 = Split(MdDNm, ".")
Select Case Sz(A1)
Case 1: Set Md = PjMd(CurPj, MdDNm)
Case 2: Set Md = PjMd(Pj(A1(0)), A1(1))
Case Else: Stop
End Select
End Function

Function MdLines$(A As CodeModule)
If A.CountOfLines = 0 Then Exit Function
MdLines = A.Lines(1, A.CountOfLines)
End Function

Function MdLy(A As CodeModule) As String()
MdLy = Split(MdLines(A), vbCrLf)
End Function

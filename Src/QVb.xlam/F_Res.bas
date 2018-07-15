Attribute VB_Name = "F_Res"
Option Explicit
Private Type Res
    Md As CodeModule
    Nm As String
End Type

Property Get ResNm_Ly(A$) As String()
'ResNm is "Pj.Md.Nm" where Pj & Md are optional
Dim O$()
Dim MthLy$()
Dim Res As Res
Res = ZResNm_Res(A)
MthLy = ZRes_MthLy(Res)
'O
    Dim J%, U%
    U = UB(MthLy)
    ReDim O(U - 2)
    For J = 1 To U - 1
        O(J - 1) = Mid(MthLy(J), 2)
    Next
ResNm_Ly = O
End Property

Private Property Get ZCurPjMd(MdNm) As CodeModule
Stop 'Set ZCurPjMd = ZCurPj.VBComponents(MdNm).CodeModule
End Property

Private Property Get ZCurVbe() As VBE
Set ZCurVbe = Excel.Application.VBE
End Property

Private Property Get ZMd(A) As CodeModule
Dim A1$(): A1 = Split(A, ".")
Select Case Sz(A1)
Case 1: Set ZMd = ZCurPjMd(A)
Case 2: Set ZMd = ZPjMd(A1(0), A1(1))
Case Else: Stop
End Select
End Property

Private Property Get ZMdLines$(A As CodeModule)
If A.CountOfLines = 0 Then Exit Property
ZMdLines = A.Lines(1, A.CountOfLines)
End Property

Private Property Get ZMdLy(A As CodeModule) As String()
ZMdLy = Split(ZMdLines(A), vbCrLf)
End Property

Private Function ZPj(A) As VBProject
Set ZPj = ZCurVbe.VBProjects(A)
End Function

Private Function ZPjMd(A, MdNm) As CodeModule
Set ZPjMd = ZPj(A).VBComponents(MdNm).CodeModule
End Function

Private Property Get ZResNm_Res(A$) As Res
Dim A1$(): A1 = Split(A, ".")
Dim O As Res
Select Case Sz(A1)
Case 1: Stop 'Set O.Md = ZCurMd:                 O.Nm = A1(0)
Case 2: Set O.Md = ZMd(A1(0)):              O.Nm = A1(1)
Case 3: Set O.Md = ZMd(A1(0) & "." & A1(1)): O.Nm = A1(2)
Case Else: Stop
End Select
ZResNm_Res = O
End Property

Private Property Get ZRes_MthLy(A As Res) As String()
Dim O$()
Dim J%, M As CodeModule, L$
Dim B$, BLno%, N%

Set M = A.Md
'N
    N = M.CountOfLines
    If N = 0 Then Exit Property
'B
    B = "Private Sub ZZRes_" & A.Nm & "()"
'BLno%: B N
    For J = M.CountOfDeclarationLines + 1 To N
        L = M.Lines(J, 1)
        If L = B Then BLno = J: Exit For
    Next
    If BLno = 0 Then Stop
'O: BLno N
    For J = BLno To N
        L = M.Lines(J, 1)
        Push O, L
        If L = "End Sub" Then Exit For
    Next
ZRes_MthLy = O
End Property

Private Sub ZZRes_XX()
'A
'B
'C
End Sub

Private Sub ZZ_ResNm_Ly()
Dim A$()
A = ResNm_Ly("XX"):           GoSub Tst
A = ResNm_Ly("F_Res.XX"):     GoSub Tst
A = ResNm_Ly("QVb.F_Res.XX"): GoSub Tst
Exit Sub
Tst:
Ass Sz(A) = 3
Ass A(0) = "A"
Ass A(1) = "B"
Ass A(2) = "C"
Return
End Sub

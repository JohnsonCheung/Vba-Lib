Attribute VB_Name = "M_Is"
Option Explicit
Function IsEmpMd(A As CodeModule) As Boolean
IsEmpMd = A.CountOfLines = 0
End Function
Function IsMdRRCCOutSideMd(MdRRCC As RRCC, Md As CodeModule) As Boolean
IsMdRRCCOutSideMd = True
Dim R%
R = MdNLin(Md)
Stop '
'If RRCC_IsEmp(MdRRCC) Then Exit Function
'With MdRRCC
'   If .R1 > R Then Exit Function
'   If .R2 > R Then Exit Function
'   If .C1 > Len(Md.Lines(.R1, 1)) + 1 Then Exit Function
'   If .C2 > Len(Md.Lines(.R2, 1)) + 1 Then Exit Function
'End With
'IsMdRRCCOutSideMd = False
End Function
Function IsOnlyTwoCdPne() As Boolean
IsOnlyTwoCdPne = CurVbe.CodePanes.Count = 2
End Function
Function IsTstMthNm(MthNm$) As Boolean
IsTstMthNm = HasSfx(MthNm, "__Tst")
End Function
Property Get IsTy() As Boolean
IsTy = HasPfx(SrcLin_RmvMdy(A), C_Ty)
End Property

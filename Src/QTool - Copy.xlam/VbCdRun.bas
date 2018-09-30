Attribute VB_Name = "VbCdRun"
Option Explicit
Private Const CdRunMdNm$ = "ZZCdRun"
Sub CdRun(Cd)
'CdAddMth Cd
Run CdAddMth(Cd)
End Sub

Private Function CdRunMd() As CodeModule
Set CdRunMd = Md(CdRunMdNm)
End Function

Function CdAddMth$(Cd)
Dim T$
T = "ZZZ" & TmpNm
MdAddSub CdRunMd, T, Cd
CdAddMth = T
End Function

Sub CdAyGen(CdAy$(), SubNm$)
MdAddSub CdRunMd, SubNm, JnCrLf(CdAy)
End Sub

Sub CdAyRun(CdAy$())
'Run CdAddMth(JnCrLf(CdAy))
CdAddMth JnCrLf(CdAy)
End Sub


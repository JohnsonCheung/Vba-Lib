Attribute VB_Name = "IdeSrt"
Option Explicit
Sub Srt_Md(Optional MdNm$)
MdSrt DftMd(MdNm)
End Sub
Function CurMdSrtedLines$()
CurMdSrtedLines = MdSrtedLines(CurMd)
End Function

Function SrcSrtedLines$(A$())
SrcSrtedLines = JnCrLf(SrcSrtedLy(A))
End Function

Function SrcSrtedLy(A$()) As String()
Dim A1$(), A2$()
A1 = SrcDclLy(A)
A2 = SrcSrtedBdyLy(A)
SrcSrtedLy = AyAddAp(A1, A2)
End Function

Function SrcSrtRpt(A$(), PjNm$, MdNm$) As DCRslt
Dim B$(): B = SrcSrtedLy(A)
Dim A1 As Dictionary
Dim B1 As Dictionary
Set A1 = SrcMthKeyLinesDic(A, PjNm, MdNm)
Set B1 = SrcMthKeyLinesDic(B, PjNm, MdNm)
Dim O As DCRslt: O = DicCmp(A1, B1, "BefSrt", "AftSrt")
SrcSrtRpt = O
End Function
Function SrcSrtRptFmt(A$(), PjNm$, MdNm$) As String()
SrcSrtRptFmt = DCRsltFmt(SrcSrtRpt(A, PjNm, MdNm))
End Function
Function SrcSrtedBdyLines$(A$())
If Sz(A) = 0 Then Exit Function
Dim D As Dictionary
Dim D1 As Dictionary
    Set D = SrcMthKeyLinesDic(A, ExlDcl:=True)
    Set D1 = DicSrt(D)
Dim O$()
    Dim K
    For Each K In D1.Keys
        Push O, vbCrLf & D1(K)
    Next
SrcSrtedBdyLines = JnCrLf(O)
End Function
Function SrcSrtedBdyLy(A$())
SrcSrtedBdyLy = SplitCrLf(SrcSrtedBdyLines(A))
End Function

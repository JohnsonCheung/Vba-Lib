Attribute VB_Name = "IdeSrt"
Option Explicit
Sub SrtMd(Optional MdNm$)
MdSrt DftMd(MdNm)
End Sub
Function CurMdSrtedLines$()
CurMdSrtedLines = MdSrtedLines(CurMd)
End Function

Function MdSrtedLines$(A As CodeModule)
MdSrtedLines = SrcSrtedLines(MdSrc(A))
End Function

Function MdSrtedLy(A As CodeModule) As String()
MdSrtedLy = SrcSrtedLy(MdSrc(A))
End Function

Function SrcSrtedLines$(A$())
If Sz(A) = 0 Then Exit Function
Dim D As Dictionary
    Set D = DicSrt(SrcSrtDic(A))
Dim O$()
    Dim K
    For Each K In D.Keys
        PushI O, D(K)
    Next
SrcSrtedLines = Join(O, vbCrLf & vbCrLf)
End Function

Function SrcSrtedLy(A$()) As String()
SrcSrtedLy = SplitCrLf(SrcSrtedLy(A))
End Function

Function SrcSrtDic(A$()) As Dictionary
Dim D As Dictionary, K
Set D = SrcDic(A)
Dim O As New Dictionary
    For Each K In D
        O.Add MthNmSrtKey(K), D(K)
    Next
Set SrcSrtDic = O
End Function

Function SrcDic(A$()) As Dictionary
Dim Ix, O As New Dictionary
O.Add "*Dcl", SrcDclLines(A)
For Each Ix In AyNz(SrcMthIx(A))
    O.Add LinMthNm(A(Ix)), SrcMthIxLines(A, Ix)
Next
Set SrcDic = O
End Function



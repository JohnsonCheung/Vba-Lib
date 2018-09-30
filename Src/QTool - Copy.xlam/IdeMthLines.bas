Attribute VB_Name = "IdeMthLines"
Option Explicit
Sub ZZ_DMthLines()
Debug.Print DMthLines("QLib.IdeInf.CmdBtnOfTileV")
End Sub
Function DMthLines$(MthDNm$)
DMthLines = MthLines(DMth(MthDNm))
End Function
Function MthLines$(A As Mth)
MthLines = SrcMthLines(MdBdyLy(A.Md), A.Nm)
End Function

Function MdMthLinesWithRmk$(A As CodeModule, MthNm)
MdMthLinesWithRmk = SrcMthLinesWithRmk(MdBdyLy(A), MthNm)
End Function

Function MthLinesWithRmk$(A As Mth)
MthLinesWithRmk = SrcMthLinesWithRmk(MdBdyLy(A.Md), A.Nm)
End Function

Function MdMthLines$(A As CodeModule, M$)
MdMthLines = MthLines(Mth(A, M))
End Function

Private Sub ZZ_MdDic()
Dim D As Dictionary
Set D = MdDic(CurMd)
Stop
End Sub

Private Sub ZZ_SrcDic()
DicBrw SrcDic(CurSrc)
End Sub



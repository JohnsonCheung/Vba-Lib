Attribute VB_Name = "IdeMthLines"
Option Explicit
Sub ZZ_DMthLines()
Debug.Print DMthLines("QLib.IdeInf.CmdBtnOfTileV")
End Sub
Function DMthLines$(MthDNm$)
DMthLines = MthLines(DMth(MthDNm))
End Function







Private Sub ZZ_SrcMthDic()
DicBrw SrcMthDic(CurSrc)
End Sub



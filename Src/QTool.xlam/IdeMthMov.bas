Attribute VB_Name = "IdeMthMov"
Option Explicit

Sub MovMth(MthPatn$, ToMdNm$)
CurMdMovMth MthPatn, Md(ToMdNm)
End Sub
Sub CurMdMovMth(MthPatn$, ToMd As CodeModule)
MdMovMth CurMd, MthPatn, ToMd
End Sub



Sub CurMthMov(ToMd$)
MthMov CurMth, Md(ToMd)
End Sub

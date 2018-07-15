VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CmdLis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Sub MdMth(Optional MthNmPatn$ = ".")
AyDmp AyAddPfx(MdMthNy(CurMd, MthNmPatn), CurMdNm & ".")
End Sub

Sub PjMth(Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".")
AyDmp CurPj.MthNy(, MthNmPatn:=MthNmPatn, MdNmPatn:=MdNmPatn)
End Sub

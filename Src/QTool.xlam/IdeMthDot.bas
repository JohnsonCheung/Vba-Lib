Attribute VB_Name = "IdeMthDot"
Option Explicit
Function CurMdMthDot(Optional WhMdyAy, Optional WhKdAy) As String()
Stop '
'CurMdMthDot = MdMthDot(CurMd, WhMdyA, WhKdAy)
End Function


Function CurPjMthDot(Optional MdRe As RegExp, Optional ExlMd$, Optional WhMdyAy, Optional WhMthKd0$) As String()
Stop '
'CurPjMthDot = PjMthDot(CurPj, MdRe, ExlMd, WhMdyA, WhMthKd0)
End Function
Function CurVbeMthDot(Optional MthRe As RegExp, Optional MthExlAy$, Optional WhMdyAy, Optional WhMthKd0$, Optional PjRe As RegExp, Optional PjExlAy$, Optional MdRe As RegExp, Optional MdExlAy$)
Stop '
'CurVbeMthDot = VbeMthDot(CurVbe, MthPatn, MthExlAy, WhMdyA, WhMthKd0, PjPatn, PjExlAy, MdPatn, MdExlAy)
End Function


Function PjMthDot(A As VBProject, Optional MthRe As RegExp, Optional MthExlAy$, Optional MdRe As RegExp, Optional MdExlAy$, Optional WhMdyAy, Optional WhMthKd0$) As String()
Stop '
'Dim MdAy(), O$(), M$(), PNm$, I, Md As CodeModule
'PNm = A.Name & "."
'For Each I In AyNz(PjStdMdAy(A, MdPatn, MdExlAy))
'    Set Md = I
'    M = SrcMthDot(MdBdyLy(Md), MthPatn, MthExlAy, WhMdyA, WhMthKd0)
'    M = AyAddPfx(M, PNm & MdNm(Md) & ".")
'    PushAy O, M
'Next
'PjMthDot = O
End Function

Function LinMthDot$(A)
LinMthDot = MthBrkDot(LinMthBrk(A))
End Function

Sub ZZ_SrcMthDot()
Brw SrcMthDot(MdBdyLy(Md("A_IdeTool")))
End Sub




Attribute VB_Name = "IdeLin"
Option Explicit

Function CurMdMthDotAy(Optional WhMdy0$, Optional WhTy0$) As String()
CurMdMthDotAy = MdMthDotAy(CurMd, WhMdy0, WhTy0)
End Function

Function MdMthDotAy(A As CodeModule, Optional WhMdy0$, Optional WhTy0$) As String()
MdMthDotAy = SrcMthDotAy(MdBdyLy(A), WhMdy0, WhTy0)
End Function

Function CurPjMthDotAy(Optional MdPatn$ = ".", Optional ExlMdLikAy0$, Optional WhMdy0$, Optional WhMthTy0$) As String()
CurPjMthDotAy = PjMthDotAy(CurPj, MdPatn, ExlMdLikAy0, WhMdy0, WhMthTy0)
End Function

Function PjMthDotAy(A As VBProject, Optional MdPatn$ = ".", Optional ExlMdLikAy0$, Optional WhMdy0$, Optional WhMthTy0$) As String()
Dim MdAy(), O$(), M$(), PNm$, I, Md As CodeModule
PNm = A.Name & "."
For Each I In AyNz(PjMdNy(A, MdPatn, ExlMdLikAy0))
    Set Md = PjMd(A, I)
    M = SrcMthDotAy(MdBdyLy(Md), WhMdy0, WhMthTy0)
    M = AyAddPfx(M, PNm & MdNm(Md) & ".")
    PushAy O, M
Next
PjMthDotAy = O
End Function

Function CmpShtToTy(Sht) As vbext_ComponentType
Select Case Sht
Case "Doc": CmpShtToTy = vbext_ComponentType.vbext_ct_Document
Case "Cls": CmpShtToTy = vbext_ComponentType.vbext_ct_ClassModule
Case "Mod": CmpShtToTy = vbext_ComponentType.vbext_ct_StdModule
Case "Frm": CmpShtToTy = vbext_ComponentType.vbext_ct_MSForm
Case "ActX": CmpShtToTy = vbext_ComponentType.vbext_ct_ActiveXDesigner
Case Else: Stop
End Select
End Function

Function CmpTyToSht$(A As vbext_ComponentType)
Select Case A
Case vbext_ComponentType.vbext_ct_Document:    CmpTyToSht = "Doc"
Case vbext_ComponentType.vbext_ct_ClassModule: CmpTyToSht = "Cls"
Case vbext_ComponentType.vbext_ct_StdModule:   CmpTyToSht = "Md"
Case vbext_ComponentType.vbext_ct_MSForm:      CmpTyToSht = "Frm"
Case vbext_ComponentType.vbext_ct_ActiveXDesigner: CmpTyToSht = "ActX"
Case Else: Stop
End Select
End Function

Function PjMdMthDotAy(A As VBProject) As String()

End Function
Function LinMthBrk(A) As String()
LinMthBrk = ShiftMthBrk(A)(0)
End Function
Function ShiftMthBrk(A) As Variant()
ReDim B$(2)
Dim L$
AyAsg ShiftShtMdy(A), B(0), L
AyAsg ShiftMthShtTy(L), B(1), L: If B(1) = "" Then ShiftMthBrk = Array(ApSy("", "", ""), A): Exit Function
AyAsg ShiftNm(L), B(2), L
ShiftMthBrk = Array(B, L)
End Function
Function MthDNmLines$(MthDNm$)
MthDNmLines = MthLines(DNmMth(MthDNm))
End Function
Function MdMthLines$(A As CodeModule, M$)
MdMthLines = MthLines(Mth(A, M))
End Function
Function ShtMdy$(Mdy)
Select Case Mdy
Case "Private": ShtMdy = "Prv"
Case "Friend": ShtMdy = "Frd"
Case "Public": ShtMdy = "Pub"
End Select
End Function

Function MthShtTy$(MthTy$)
Select Case MthTy
Case "Property Get": MthShtTy = "Get"
Case "Property Set": MthShtTy = "Set"
Case "Property Let": MthShtTy = "Get"
Case "Function": MthShtTy = "Fun"
Case "Sub": MthShtTy = "Get"
End Select
End Function

Function TakMthTy$(A)
TakMthTy = TakPfxAy(A, MthTyAy)
End Function
Function TakMthShtTy$(A)
Dim B$
B = TakPfxAy(A, MthTyAy): If B = "" Then Exit Function
TakMthShtTy = MthShtTy(B)
End Function

Function MthBrkDot$(MthBrk$())
If MthBrk(2) = "" Then Exit Function
MthBrkDot = JnDot(MthBrk)
End Function

Function LinMthDot$(A)
LinMthDot = MthBrkDot(LinMthBrk(A))
End Function

Function JnDot$(A)
JnDot = Join(A, ".")
End Function

Attribute VB_Name = "IdeLin"
Option Explicit

Function CmpShtToTy(Sht) As vbext_ComponentType
Select Case Sht
Case "Doc": CmpShtToTy = vbext_ComponentType.vbext_ct_Document
Case "Cls": CmpShtToTy = vbext_ComponentType.vbext_ct_ClassModule
Case "Std": CmpShtToTy = vbext_ComponentType.vbext_ct_StdModule
Case "Frm": CmpShtToTy = vbext_ComponentType.vbext_ct_MSForm
Case "ActX": CmpShtToTy = vbext_ComponentType.vbext_ct_ActiveXDesigner
Case Else: Stop
End Select
End Function

Function CmpTyToSht$(A As vbext_ComponentType)
Select Case A
Case vbext_ComponentType.vbext_ct_Document:    CmpTyToSht = "Doc"
Case vbext_ComponentType.vbext_ct_ClassModule: CmpTyToSht = "Cls"
Case vbext_ComponentType.vbext_ct_StdModule:   CmpTyToSht = "Std"
Case vbext_ComponentType.vbext_ct_MSForm:      CmpTyToSht = "Frm"
Case vbext_ComponentType.vbext_ct_ActiveXDesigner: CmpTyToSht = "ActX"
Case Else: Stop
End Select
End Function


Function MthDotNTM$(MthDot$)
'MthDot is a string with last 3 seg as Mdy.ShtTy.Nm
'MthNTM is a string with last 3 seg as Nm:ShtTy.Mdy
Dim Ay$(), Nm$, ShtTy$, Mdy$
Ay = SplitDot(MthDot)
AyAsg AyPop(Ay), Ay, Nm
AyAsg AyPop(Ay), Ay, ShtTy
AyAsg AyPop(Ay), Ay, Mdy
Push Ay, FmtQQ("?:?.?", Nm, ShtTy, Mdy)
MthDotNTM = JnDot(Ay)
End Function


Function MthDNmLines$(MthDNm$)
MthDNmLines = MthLines(DMth(MthDNm))
End Function
Function ShtMdy$(Mdy)
Select Case Mdy
Case "": Exit Function
Case "Private": ShtMdy = "Prv"
Case "Friend": ShtMdy = "Frd"
Case "Public": ShtMdy = "Pub"
End Select
End Function

Function MthShtTy$(MthTy)
Dim O$
Select Case MthTy
Case "Sub": O = MthTy
Case "Function": O = "Fun"
Case "Property Get": O = "Get"
Case "Property Let": O = "Let"
Case "Property Set": O = "Set"
End Select
MthShtTy = O
End Function

Function TakMthTy$(A)
TakMthTy = TakPfxAy(A, MthTyAy)
End Function
Function TakMthKd$(A)
TakMthKd = TakPfxAyS(A, MthKdAy)
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


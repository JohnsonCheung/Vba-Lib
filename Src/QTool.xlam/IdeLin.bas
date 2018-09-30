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





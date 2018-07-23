Attribute VB_Name = "G_Ide"
Public Const TyChrLis$ = "!@#$%^&"
Public Enum eTstLABCs
    eValidateAsFldVal
    eValidateAsNm
    eValidateAsFny
    eValidateAsBetNum
    eAll
End Enum

'Property Get Fxa(A) As Fxa
'Dim O As New Fxa
'Set Fxa = O.Init(A)
'End Property
Function CmpTy_Str$(A As vbext_ComponentType)
Dim O$
Select Case A
Case vbext_ComponentType.vbext_ct_ActiveXDesigner: O = "ActiveXDesigner"
Case vbext_ComponentType.vbext_ct_ClassModule: O = "Class"
Case vbext_ComponentType.vbext_ct_Document: O = "Doc"
Case vbext_ComponentType.vbext_ct_MSForm: O = "MsForm"
Case vbext_ComponentType.vbext_ct_StdModule: O = "Md"
Case Else: O = "Unknown(" & A & ")"
End Select
CmpTy_Str = O
End Function

Function DftVbe(A As VBE) As VBE
If IsNothing(A) Then
   Set DftVbe = CurVbe
Else
   Set DftVbe = A
End If
End Function

Function IsTyChr(S) As Boolean
If Len(S) <> 1 Then Exit Function
IsTyChr = HasSubStr(TyChrLis, S)
End Function

Function MthDotNm_Mth(A$) As Mth
Dim M As CodeModule
Dim Nm$
Dim Ny$(): Ny = Split(A, ".")
Select Case Sz(Ny)
Case 1: Nm = Ny(0): Set M = CurMd
Case 2: Nm = Ny(1): Set M = Md(Ny(0))
Case 3: Nm = Ny(2): Set M = PjMd(Pj(Ny(0)), Ny(1))
Case Else: Stop
End Select
Set MdMthDotNm_Brk = Mth(M, Nm)
End Function

Function RfPth$(A As VBIDE.Reference)
On Error Resume Next
RfPth = A.FullPath
End Function

Function RfToStr$(A As VBIDE.Reference)
With A
   RfToStr = .Name & " " & RfPth(A)
End With
End Function

Function TyChrAsTyStr$(TyChr$)
Dim O$
Select Case TyChr
Case "#": O = "Double"
Case "%": O = "Integer"
Case "!": O = "Signle"
Case "@": O = "Currency"
Case "^": O = "LongLong"
Case "$": O = "String"
Case "&": O = "Long"
Case Else: Stop
End Select
TyChrAsTyStr = O
End Function

Function VbeCmdBarAy(A As VBE) As Office.CommandBar()
Dim O() As Office.CommandBar
Dim I
For Each I In A.CommandBars
   PushObj O, I
Next
VbeCmdBarAy = O
End Function

Function VbeCmdBarNy(A As VBE) As String()
Stop '
'VbeCmdBarNy = ItrNy(A.CommandBars)
End Function

Function WinAy() As VBIDE.Window()
Dim O() As VBIDE.Window, W As VBIDE.Window
Stop '
'For Each W In Vbe.Windows
'   PushObj O, W
'Next
WinAy = O
End Function

Function WinAyOfCd() As VBIDE.Window()
WinAyOfCd = WinAyOfTy(vbext_wt_CodeWindow)
End Function

Function WinAyOfTy(T As vbext_WindowType) As VBIDE.Window()
WinAyOfTy = OyWhPrp(WinAy, "Type", T)
End Function

Function WinCnt&()
WinCnt = Application.VBE.Windows.Count
End Function

Function WinMdNm$(A As VBIDE.Window)
WinMdNm = TakBet(A.Caption, " - ", " (Code)")
End Function

Property Get CurCdPne() As VBIDE.CodePane
Set CurCdPne = CurVbe.ActiveCodePane
End Property

Sub SrcPth_BldFxa(SrcPth$)
Stop '
Dim F$
   Dim Fnn$
'   Fnn = FfnFnn(RmvLasChr(SrcPth))
   F = SrcPth & Fnn & ".xlam"
Dim X As Excel.Application
'   Set X = Fxa(F).Crt
Dim P As VBProject
'Set P = Pjx(X.Vbe.VBProjects(1))

Dim SrcFfnAy$()
   Dim S
   Stop '
'   SrcFfnAy = AyWhLikAy(PthFfnAy(SrcPth), LvsSy("*.bas *.cls"))
   For Each S In SrcFfnAy
       P.ImpSrcFfn S
   Next
   P.RmvOptCmpDbLin
   P.ImpRf SrcPth
Dim Wb As Workbook
   Set Wb = X.Workbooks(1)
Wb.SaveAs F, FileFormat:=XlFileFormat.xlOpenXMLAddIn
Wb.Close
Wb.Quit
Set X = Nothing
End Sub

Sub WinClsAll()
Dim W As VBIDE.Window
For Each W In Application.VBE.Windows
   W.Close
Next
End Sub

Sub WinClsCd(Optional ExceptMdNm$)
Dim I, W As VBIDE.Window
For Each I In WinAyOfCd
   Set W = I
   If WinMdNm(W) <> ExceptMdNm Then
       W.Close
   End If
Next
End Sub

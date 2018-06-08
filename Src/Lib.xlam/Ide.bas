Attribute VB_Name = "Ide"
Option Explicit
Public Const TyChrLis$ = "!@#$%^&"
Public Enum eTstLABCs
    eValidateAsFldVal
    eValidateAsNm
    eValidateAsFny
    eValidateAsBetNum
    eAll
End Enum
Property Get CvPj(I) As Pjx
Set CvPj = I
End Property
Property Get CvPjx(I) As Pjx
Set CvPjx = I
End Property

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

Property Get Fxa(A) As Fxa
Dim O As New Fxa
Set Fxa = O.Init(A)
End Property
Property Get Md(MdNm) As CodeModule
Dim A As VBComponents: Set A = CurPj.VBComponents
Dim I, Cmp As VBComponent
Set Md = CurPj.VBComponents(MdNm).CodeModule
End Property

Property Get SrcLin(A) As SrcLin
Dim O As New SrcLin
Set SrcLin = O.Init(A)
End Property

Property Get CurCdPne() As VBIde.CodePane
Set CurCdPne = CurVbe.ActiveCodePane
End Property

Property Get CurCdWin() As VBIde.Window
Set CurCdWin = Vbe.ActiveCodePane.Window
End Property

Property Get CurVbe() As Vbe
Set CurVbe = Application.Vbe
End Property

Property Get CurWin() As VBIde.Window
Set CurWin = CurVbe.ActiveWindow
End Property

Property Get Dcl(A$()) As Dcl
Dim O As New Dcl
Set Dcl = O.Init(A)
End Property
Function DftVbe(A As Vbe) As Vbe
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

Function RfPth$(A As VBIde.Reference)
On Error Resume Next
RfPth = A.FullPath
End Function

Function RfToStr$(A As VBIde.Reference)
With A
   RfToStr = .Name & " " & RfPth(A)
End With
End Function

Sub SrcPth_BldFxa(SrcPth$)
Dim F$
   Dim Fnn$
   Fnn = FfnFnn(RmvLasChr(SrcPth))
   F = SrcPth & Fnn & ".xlam"
Dim X As Excel.Application
   Set X = Fxa(F).Crt
Dim P As Pjx
Set P = Pjx(X.Vbe.VBProjects(1))

Dim SrcFfnAy$()
   Dim S
   SrcFfnAy = AyWhLikAy(PthFfnAy(SrcPth), LvsSy("*.bas *.cls"))
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

Function VbeCmdBarAy(A As Vbe) As Office.CommandBar()
Dim O() As Office.CommandBar
Dim I
For Each I In A.CommandBars
   PushObj O, I
Next
VbeCmdBarAy = O
End Function

Function VbeCmdBarNy(A As Vbe) As String()
VbeCmdBarNy = ItrNy(A.CommandBars)
End Function

Function WinAy() As VBIde.Window()
Dim O() As VBIde.Window, W As VBIde.Window
For Each W In Vbe.Windows
   PushObj O, W
Next
WinAy = O
End Function

Function WinAyOfCd() As VBIde.Window()
WinAyOfCd = WinAyOfTy(vbext_wt_CodeWindow)
End Function

Function WinAyOfTy(T As vbext_WindowType) As VBIde.Window()
WinAyOfTy = Oy(WinAy).WhPrp("Type", T)
End Function

Sub WinClsAll()
Dim W As VBIde.Window
For Each W In Application.Vbe.Windows
   W.Close
Next
End Sub

Sub WinClsCd(Optional ExceptMdNm$)
Dim I, W As VBIde.Window
For Each I In WinAyOfCd
   Set W = I
   If WinMdNm(W) <> ExceptMdNm Then
       W.Close
   End If
Next
End Sub

Function WinCnt&()
WinCnt = Application.Vbe.Windows.Count
End Function
Property Get Vbex(A As Vbe) As Vbex
Dim O As New Vbex
Set Vbex = O.Init(A)
End Property
Function WinMdNm$(A As VBIde.Window)
WinMdNm = TakBet(A.Caption, " - ", " (Code)")
End Function
Sub TstSrcLin()
Dim A As New SrcLin: A.Tst
End Sub
Function CvMd(A) As CodeModule
Set CvMd = A
End Function
Function CvMdx(A) As Mdx
CvMdx = A
End Function
Property Get Mdx(A As CodeModule) As Mdx
Dim O As New Mdx
Set Mdx = O.Init(A)
End Property
Property Get Pjx(A) As Pjx
Dim O As New Pjx
Set Pjx = O.Init(A)
End Property
Property Get CurPj() As VBProject
Set CurPj = CurVbe.ActiveVBProject
End Property
Property Get CurPjx() As Pjx
Set CurPjx = Pjx(CurPj)
End Property

Property Get CurVbex() As Vbex
Set CurVbex = Vbex(CurVbe)
End Property

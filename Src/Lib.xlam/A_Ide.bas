Attribute VB_Name = "A_Ide"
Option Explicit
Public Const TyChrLis$ = "!@#$%^&"
Property Get Ide() As Ide
Static Y As New Ide
Set Ide = Y
End Property
Function CurCdPne() As VBIDE.CodePane
Set CurCdPne = CurVbe.ActiveCodePane
End Function

Function CurCdWin() As VBIDE.Window
Set CurCdWin = VBE.ActiveCodePane.Window
End Function

Function CurVbe() As VBE
Set CurVbe = Application.VBE
End Function

Function CurWin() As VBIDE.Window
Set CurWin = CurVbe.ActiveWindow
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


Function RfPth$(A As VBIDE.Reference)
On Error Resume Next
RfPth = A.FullPath
End Function

Function RfToStr$(A As VBIDE.Reference)
With A
   RfToStr = .Name & " " & RfPth(A)
End With
End Function

Sub SrcPth_BldFxa(SrcPth$)
Dim Fxa$
   Dim Fnn$
   Fnn = FfnFnn(RmvLasChr(SrcPth))
   Fxa = SrcPth & Fnn & ".xlam"
Dim X As Excel.Application
   Set X = FxaCrt(Fxa)
Dim Pj As Pj
Set Pj = A_Ide.Pj(X.VBE.VBProjects(1))

Dim SrcFfnAy$()
   Dim S
   SrcFfnAy = AyWhLikAy(PthFfnAy(SrcPth), LvsSy("*.bas *.cls"))
   For Each S In SrcFfnAy
       Pj.ImpSrcFfn S
   Next
   Pj.RmvOptCmpDbLin
   Pj.ImpRf SrcPth
Dim Wb As Workbook
   Set Wb = X.Workbooks(1)
Wb.SaveAs Fxa, FileFormat:=XlFileFormat.xlOpenXMLAddIn
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

Function VbeCmdBarAy(A As VBE) As Office.CommandBar()
Dim O() As Office.CommandBar
Dim I
For Each I In A.CommandBars
   PushObj O, I
Next
VbeCmdBarAy = O
End Function

Function VbeCmdBarNy(A As VBE) As String()
VbeCmdBarNy = ItrNy(A.CommandBars)
End Function

Function WinAy() As VBIDE.Window()
Dim O() As VBIDE.Window, W As VBIDE.Window
For Each W In VBE.Windows
   PushObj O, W
Next
WinAy = O
End Function

Function WinAyOfCd() As VBIDE.Window()
WinAyOfCd = WinAyOfTy(vbext_wt_CodeWindow)
End Function

Function WinAyOfTy(T As vbext_WindowType) As VBIDE.Window()
WinAyOfTy = OyWhPrp(WinAy, "Type", T)
End Function

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

Function WinCnt&()
WinCnt = Application.VBE.Windows.Count
End Function

Sub TstOy()
Dim O As New Oy: O.Tst
End Sub

Sub AAA()
TstP3
End Sub

Sub TstLoFmtr()
Dim O As New LoFmtr: O.Tst
End Sub

Sub TstP3()
Dim O As New P3: O.Tst

End Sub

Sub TstP3LCFVRslt()
Dim O As New P3LCFVRslt: O.Tst
End Sub

Function WinMdNm$(A As VBIDE.Window)
WinMdNm = TakBet(A.Caption, " - ", " (Code)")
End Function

Function CvMd(A) As CodeModule
Set CvMd = A
End Function
Function CvMdx(A) As Md
CvMdx = A
End Function
Function CurPj() As VBProject
Set CurPj = CurVbe.ActiveVBProject
End Function
Property Get Md(A As CodeModule) As Md
Dim O As New Md
O.Init A
Md.Nm
Set Md = O
End Property
Property Get Pj(A As VBProject) As Pj
Dim O As New Pj
Set Pj = O.Init(A)
End Property
Property Get CurPjx() As Pj
Set CurPjx = Pj(CurPj)
End Property

Attribute VB_Name = "Ide"
Option Explicit
Public Const TyChrLis$ = "!@#$%^&"

Function WinAy() As VBIDE.Window()
Dim O() As VBIDE.Window, W As VBIDE.Window
For Each W In VBE.Windows
   PushObj O, W
Next
WinAy = O
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

Function CurVbe() As VBE
Set CurVbe = Application.VBE
End Function

Function CurWin() As VBIDE.Window
Set CurWin = CurVbe.ActiveWindow
End Function

Function CurCdPne() As VBIDE.CodePane
Set CurCdPne = CurVbe.ActiveCodePane
End Function

Function CurCdWin() As VBIDE.Window
Set CurCdWin = VBE.ActiveCodePane.Window
End Function

Function DftVbe(A As VBE) As VBE
If IsNothing(A) Then
   Set DftVbe = CurVbe
Else
   Set DftVbe = A
End If
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

Function WinMdNm$(A As VBIDE.Window)
WinMdNm = TakBet(A.Caption, " - ", " (Code)")
End Function

Function IsTyChr(S) As Boolean
If Len(S) <> 1 Then Exit Function
IsTyChr = HasSubStr(TyChrLis, S)
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

Function RfPth$(A As VBIDE.Reference)
On Error Resume Next
RfPth = A.FullPath
End Function

Function RfToStr$(A As VBIDE.Reference)
With A
   RfToStr = .Name & " " & RfPth(A)
End With
End Function

Function PjFilNm$(A As VBProject)
On Error Resume Next
PjFilNm = DftPj(A).Filename
End Function

Sub PjImp(A As VBProject, SrcFfn)
A.VBComponents.Import SrcFfn
End Sub

Sub PjRmvOptCmpDbLin(A As VBProject)
Dim I, Md As CodeModule
For Each I In PjMdAy(A)
   Set Md = I
   MdRmvOptCmpDb Md
Next
End Sub

Sub SrcPth_BldFxa(SrcPth$)
Dim Fxa$
   Dim Fnn$
   Fnn = FfnFnn(RmvLasChr(SrcPth))
   Fxa = SrcPth & Fnn & ".xlam"
Dim X As Excel.Application
   Set X = FxaCrt(Fxa)
Dim Pj As VBProject
Set Pj = X.VBE.VBProjects(1)

Dim SrcFfnAy$()
   Dim S
   SrcFfnAy = AyWhLikAy(PthFfnAy(SrcPth), LvsSy("*.bas *.cls"))
   For Each S In SrcFfnAy
       PjImp Pj, S
   Next
   PjRmvOptCmpDbLin Pj
   PjImpRf Pj, SrcPth
Dim Wb As Workbook
   Set Wb = X.Workbooks(1)
Wb.SaveAs Fxa, FileFormat:=XlFileFormat.xlOpenXMLAddIn
Wb.Close
Wb.Quit
Set X = Nothing
End Sub

Sub SrcPth_BldXla__Tst()
Dim SrcPth$
   SrcPth = PjSrcPth(CurPj)
SrcPth_BldFxa SrcPth
Stop
End Sub


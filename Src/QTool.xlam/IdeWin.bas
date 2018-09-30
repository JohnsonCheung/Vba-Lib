Attribute VB_Name = "IdeWin"
Option Explicit
Sub ClsAllWin()
Dim W As VBIDE.Window
For Each W In CurVbe.Windows
    W.Close
Next
End Sub
Sub ShwDbg()
Dim W As VBIDE.Window
Set W = CurMdWin
If IsNothing(W) Then
    ClsWinExl CurMdWin
    LclWin.Visible = True
    TileVBtn.Execute
End If
End Sub
Function CurMdWin() As VBIDE.Window
If IsNothing(CurVbe.ActiveCodePane) Then Exit Function
Set CurMdWin = CurVbe.ActiveCodePane.Window
End Function
Sub ClsWinExl(A As VBIDE.Window)
Dim W As VBIDE.Window
For Each W In AllWin
    If Not IsEqObj(A, W) Then
        W.Close
    End If
    W.Close
Next
WinCls A
End Sub
Function WinTy(A As VBIDE.Window) As VBIDE.vbext_WindowType
On Error GoTo X
WinTy = A.Type
Exit Function
X: WinTy = -1
End Function
Sub WinCls(A As VBIDE.Window)
If IsNothing(A) Then Exit Sub
If WinTy(A) = -1 Then Exit Sub
If A.Visible Then A.Close
End Sub

Function AllWin() As VBIDE.Windows
Set AllWin = CurVbe.Windows
End Function
Sub ClsWin()
VbeClsWin CurVbe
End Sub
Function CurCdWin() As VBIDE.Window
Dim C As VBComponent: Set C = CurCmp: If IsNothing(C) Then Exit Function
Dim M As CodeModule: Set M = C.CodeModule: If IsNothing(M) Then Exit Function
Set CurCdWin = M.CodePane.Window
End Function
Function ImmWin() As VBIDE.Window
Set ImmWin = WinTyWin(vbext_wt_Immediate)
End Function
Function WinTyWin(Ty As vbext_WindowType) As VBIDE.Window
Set WinTyWin = CurVbe.Windows(Ty)
End Function
Sub ClsWinExptImm(Optional ExcptWinTyAy)
VbeClsWin CurVbe, Array(VBIDE.vbext_wt_Immediate)
End Sub
Function LclWin() As VBIDE.Window
Set LclWin = WinTyWin(vbext_wt_Locals)
End Function

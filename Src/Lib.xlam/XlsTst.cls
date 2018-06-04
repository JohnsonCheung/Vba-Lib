VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XlsTst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Sub GenCmdTstFun()
Const Qvbl$ = "Sub ?()|Dim A As New ?: A.Tst|End Sub"
Dim Bdy$: Bdy = Seed(Qvbl).Expand(AySrt(CurPjx.ClsNy))
With CurMd
    Dim A%, N%
    A = 13
    N = .CountOfLines - A
    .DeleteLines A, N
    .AddFromString Bdy
End With
End Sub

Sub LABCs()
Dim O As New LABCs: O.Tst
End Sub

Sub DDLines()
Dim A As New DDLines: A.Brw
End Sub

Sub LoFmtr()
Dim A As New LoFmtr: A.Tst
End Sub

Sub FmtWs()
Dim A As New FmtWs: A.Tst
End Sub


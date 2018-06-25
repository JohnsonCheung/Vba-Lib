VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Sub GenCmdTstFun()
Const Qvbl$ = "Sub ?()|Dim A As New ?: A.Tst|End Sub"
Dim Bdy$: Stop:  'Bdy = Seed(Qvbl).Expand(AySrt(CurPjx.ClsNy))
Stop
'With CurMd
'    Dim A%, N%
'    A = 13
'    N = .CountOfLines - A
'    .DeleteLines A, N
'    .AddFromString Bdy
'End With
End Sub


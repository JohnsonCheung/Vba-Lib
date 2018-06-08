VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ffn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A$ 'Ffn
Friend Property Get Init(Ffn) As Ffn
A = Ffn
Set Init = Me
End Property
Function AddFnSfx$(Sfx$)
AddFnSfx = RmvExt & Sfx & Ext
End Function

Sub CpyToPth(ToPth$, Optional OvrWrt As Boolean)
Fso.CopyFile A, ToPth$ & Fn, OvrWrt
End Sub

Sub Dlt()
If IsExist Then Kill A
End Sub

Function Ext$()
Dim P%: P = InStrRev(A, ".")
If P = 0 Then Exit Function
Ext = Mid(A, P)
End Function

Function Fdr$()
Stop
'Fdr = PthFdr(Pth)
End Function

Function Fn$()
Dim P%: P = InStrRev(A, "\")
If P = 0 Then Fn = A: Exit Function
Fn = Mid(A, P + 1)
End Function

Function Fnn$()
Fnn = Vb.Ffn(Fn).RmvExt
End Function

Function IsExist() As Boolean
IsExist = Fso.FileExists(A)
End Function

Function Pth$()
Dim P%: P = InStrRev(A, "\")
If P = 0 Then Exit Function
Pth = Left(A, P)
End Function

Function RmvExt$()
Dim P%: P = InStrRev(Fn, ".")
If P = 0 Then RmvExt = Left(A, P): Exit Function
RmvExt = Left(A, P - 1)
End Function

Function RplExt$(NewExt)
RplExt = RmvExt & NewExt
End Function


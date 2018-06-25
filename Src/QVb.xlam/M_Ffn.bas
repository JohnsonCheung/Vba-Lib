Attribute VB_Name = "M_Ffn"
Option Explicit

Property Get FfnAddFnSfx$(A, Sfx$)
FfnAddFnSfx = FfnRmvExt(A) & Sfx & FfnExt(A)
End Property

Property Get FfnExt$(A)
Dim P%: P = InStrRev(A, ".")
If P = 0 Then Exit Property
FfnExt = Mid(A, P)
End Property

Property Get FfnFdr$(A)
FfnFdr = PthFdr(FfnPth(A))
End Property

Property Get FfnFn$(A)
Dim P%: P = InStrRev(A, "\")
If P = 0 Then FfnFn = A: Exit Property
FfnFn = Mid(A, P + 1)
End Property

Property Get FfnFnn$(A)
FfnFnn = FfnRmvExt(A)
End Property

Property Get FfnIsExist(A) As Boolean
FfnIsExist = Fso.FileExists(A)
End Property

Property Get FfnPth$(A)
Dim P%: P = InStrRev(A, "\")
If P = 0 Then Exit Property
FfnPth = Left(A, P)
End Property

Property Get FfnRmvExt$(A)
Dim P%: P = InStrRev(A, ".")
If P = 0 Then FfnRmvExt = Left(A, P): Exit Property
FfnRmvExt = Left(A, P - 1)
End Property

Property Get FfnRplExt$(A, NewExt)
FfnRplExt = FfnRmvExt(A) & NewExt
End Property

Sub FfnCpyToPth(A, ToPth$, Optional OvrWrt As Boolean)
Fso.CopyFile A, ToPth$ & FfnFn(A), OvrWrt
End Sub

Sub FfnDlt(A)
If FfnIsExist(A) Then Kill A
End Sub

Attribute VB_Name = "M_Ffn"
Option Explicit
Function FfnSz&(A)
If Not FfnIsExist(A) Then FfnSz = -1: Exit Function
FfnSz = FileLen(A)
End Function
Function FfnTim(A) As Date
If Not FfnIsExist(A) Then Exit Function
FfnTim = FileDateTime(A)
End Function
Function FfnAddFnSfx$(A, Sfx$)
FfnAddFnSfx = FfnRmvExt(A) & Sfx & FfnExt(A)
End Function

Function FfnExt$(A)
Dim P%: P = InStrRev(A, ".")
If P = 0 Then Exit Function
FfnExt = Mid(A, P)
End Function

Function FfnFdr$(A)
FfnFdr = PthFdr(FfnPth(A))
End Function

Function FfnFn$(A)
Dim P%: P = InStrRev(A, "\")
If P = 0 Then FfnFn = A: Exit Function
FfnFn = Mid(A, P + 1)
End Function

Function FfnFnn$(A)
FfnFnn = FfnRmvExt(A)
End Function

Function FfnIsExist(A) As Boolean
FfnIsExist = Fso.FileExists(A)
End Function

Function FfnPth$(A)
Dim P%: P = InStrRev(A, "\")
If P = 0 Then Exit Function
FfnPth = Left(A, P)
End Function

Function FfnRmvExt$(A)
Dim P%: P = InStrRev(A, ".")
If P = 0 Then FfnRmvExt = Left(A, P): Exit Function
FfnRmvExt = Left(A, P - 1)
End Function

Function FfnRplExt$(A, NewExt)
FfnRplExt = FfnRmvExt(A) & NewExt
End Function

Sub FfnCpyToPth(A, ToPth$, Optional OvrWrt As Boolean)
Fso.CopyFile A, ToPth$ & FfnFn(A), OvrWrt
End Sub

Sub FfnDlt(A)
On Error GoTo X
Kill A
Exit Sub
X: Debug.Print FmtQQ("FfnDtl: Kill(?) Er(?)", A, Err.Description)
End Sub

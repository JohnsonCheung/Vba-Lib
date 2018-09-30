Attribute VB_Name = "VbFsFfn"
Option Explicit

Function FfnTim(A) As Date
FfnTim = FileDateTime(A)
End Function

Sub FfnDlt(A)
On Error GoTo X
Kill A
Exit Sub
X: Debug.Print FmtQQ("FfnDtl: Kill(?) Er(?)", A, Err.Description)
End Sub

Function FfnAddFnSfx$(A$, Sfx$)
FfnAddFnSfx = CutExt(A) & Sfx & FfnExt(A)
End Function

Function FfnExt$(A$)
Dim B$, P%
B = LasFilSeg(A)
P = InStrRev(B, ".")
If P = 0 Then Exit Function
FfnExt = Mid(B, P)
End Function

Function FfnNxt$(A$)
If FfnIsExist(A) Then FfnNxt = A: Exit Function
Dim J%, O$
For J = 1 To 999
    O = FfnAddFnSfx(A, "(" & Format(J, "000") & ")")
    If Not FfnIsExist(O) Then FfnNxt = O: Exit Function
Next
Stop
End Function

Function FfnFn$(A)
Dim P%: P = InStrRev(A, "\")
If P = 0 Then FfnFn = A: Exit Function
FfnFn = Mid(A, P + 1)
End Function

Function FfnFnn$(A)
FfnFnn = FfnRmvExt(FfnFn(A))
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
Sub FfnCpyToPth(A, ToPth$, Optional OvrWrt As Boolean)
Fso.CopyFile A, ToPth$ & FfnFn(A), OvrWrt
End Sub

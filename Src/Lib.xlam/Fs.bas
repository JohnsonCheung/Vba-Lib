Attribute VB_Name = "Fs"
Option Explicit
Private O$() ' Used by PthEntAyR

Function DftFfn(Ffn0, Optional Ext$ = ".txt", Optional Pth0$, Optional Fdr$)
If Ffn0 <> "" Then DftFfn = Ffn0: Exit Function
Dim Pth$: Pth = DftPth(Pth0)
DftFfn = Pth & TmpNm & Ext
End Function

Function DftPth$(Optional Pth0$, Optional Fdr$)
If Pth0 <> "" Then DftPth = Pth0: Exit Function
DftPth = TmpPth(Fdr)
End Function

Function FfnAddFnSfx(A$, Sfx$)
FfnAddFnSfx = FfnRmvExt(A) & Sfx & FfnExt(A)
End Function

Sub FfnCpyToPth(A$, ToPth$, Optional OvrWrt As Boolean)
Fso.CopyFile A, ToPth$ & FfnFn(A), OvrWrt
End Sub

Sub FfnDlt(Ffn)
If FfnIsExist(Ffn) Then Kill Ffn
End Sub

Function FfnExt$(Ffn)
Dim P%: P = InStrRev(Ffn, ".")
If P = 0 Then Exit Function
FfnExt = Mid(Ffn, P)
End Function

Function FfnFdr$(Ffn)
FfnFdr = PthFdr(FfnPth(Ffn))
End Function

Function FfnFn$(Ffn)
Dim P%: P = InStrRev(Ffn, "\")
If P = 0 Then FfnFn = Ffn: Exit Function
FfnFn = Mid(Ffn, P + 1)
End Function

Function FfnFnn$(Ffn)
FfnFnn = FfnRmvExt(FfnFn(Ffn))
End Function

Function FfnIsExist(Ffn) As Boolean
FfnIsExist = Fso.FileExists(Ffn)
End Function

Function FfnPth$(Ffn)
Dim P%: P = InStrRev(Ffn, "\")
If P = 0 Then Exit Function
FfnPth = Left(Ffn, P)
End Function

Function FfnRmvExt(Fn)
Dim P%: P = InStrRev(Fn, ".")
If P = 0 Then FfnRmvExt = Left(Fn, P): Exit Function
FfnRmvExt = Left(Fn, P - 1)
End Function

Function FfnRplExt$(Ffn, NewExt)
FfnRplExt = FfnRmvExt(Ffn) & NewExt
End Function

Sub FtBrw(Ft)
'Shell "code.cmd """ & Ft & """", vbHide
Shell "notepad.exe """ & Ft & """", vbMaximizedFocus
End Sub

Function FtLines$(Ft)
FtLines = Fso.GetFile(Ft).OpenAsTextStream.ReadAll
End Function
Function FtDic(Ft) As Dictionary
Set FtDic = Ly(FtLy(Ft)).Dic
End Function

Function FtLy(Ft) As String()
Dim F%: F = FtOpnInp(Ft)
Dim L$, O$()
While Not EOF(F)
    Line Input #F, L
    Push O, L
Wend
Close #F
FtLy = O
End Function

Function FtOpnApp%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Append As #O
FtOpnApp = O
End Function

Function FtOpnInp%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Input As #O
FtOpnInp = O
End Function

Function FtOpnOup%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Output As #O
FtOpnOup = O
End Function

Sub PthBrw(P)
Shell "Explorer """ & P & """", vbMaximizedFocus
End Sub

Sub PthClrFil(A$)
If Not PthIsExist(A) Then Exit Sub
Dim Ay$(): Ay = PthFfnAy(A)
Dim F
On Error Resume Next
For Each F In Ay
   Kill F
Next
End Sub

Sub PthEns(P$)
If PthIsExist(P) Then Exit Sub
MkDir P
End Sub

Function PthEntAy(A, Optional FilSpec$ = "*.*", Optional Atr As FileAttribute, Optional IsRecursive As Boolean) As String()
If Not IsRecursive Then
    PthEntAy = AyAdd(PthSubPthAy(A), PthFfnAy(A, FilSpec, Atr))
    Exit Function
End If
Erase O
PthPushEntAyR A
PthEntAy = O
Erase O
End Function

Function PthFdr$(A$)
Ass PthHasPthSfx(A)
Dim P$: P = RmvLasChr(A)
PthFdr = TakAftRev(A, "\")
End Function

Function PthFfnAy(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
PthFfnAy = AyAddPfx(PthFnAy(A, Spec, Atr), A)
End Function

Function PthFnAy(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
Ass PthIsExist(A)
Dim O$()
Dim M$
M = Dir(A & Spec)
If Atr = 0 Then
    While M <> ""
       Push O, M
       M = Dir
    Wend
    PthFnAy = O
End If
Ass PthHasPthSfx(A)
While M <> ""
    If GetAttr(A & M) And Atr Then
        Push O, M
    End If
    M = Dir
Wend
PthFnAy = O
End Function

Function PthHasFil(A) As Boolean
Ass PthHasPthSfx(A)
If Not PthIsExist(A) Then Exit Function
PthHasFil = (Dir(A & "*.*") <> "")
End Function

Function PthHasPthSfx(A) As Boolean
PthHasPthSfx = LasChr(A) = "\"
End Function

Function PthHasSubDir(A) As Boolean
If Not PthIsExist(A) Then Exit Function
Ass PthHasPthSfx(A)
Dim P$: P = Dir(A & "*.*", vbDirectory)
Dir
PthHasSubDir = Dir <> ""
End Function

Function PthIsEmp(A)
Ass PthIsExist(A)
If PthHasFil(A) Then Exit Function
If PthHasSubDir(A) Then Exit Function
PthIsEmp = True
End Function

Function PthIsExist(A) As Boolean
Ass PthHasPthSfx(A)
PthIsExist = Fso.FolderExists(A)
End Function

Sub PthRmvEmpSubDir(A)
Dim P$(): P = PthSubPthAy(A): If AyIsEmp(A) Then Exit Sub
Dim I
For Each I In P
   PthRmvIfEmp CStr(I)
Next
End Sub

Sub PthRmvIfEmp(A$)
If Not PthIsExist(A) Then Exit Sub
If PthIsEmp(A) Then Exit Sub
RmDir A
End Sub

Function PthSubFdrAy(A, Optional Spec$ = "*.*", Optional Atr As VbFileAttribute) As String()
'PthSubFdrAy = ItrNy(Fso.GetFolder(A).SubFolders, Spec)
Ass PthIsExist(A)
Ass PthHasPthSfx(A)
Dim O$(), M$, X&, XX&
X = Atr Or vbDirectory
M = Dir(A & Spec, vbDirectory)
While M <> ""
    If InStr(M, "?") > 0 Then
        Debug.Print "PthSubFdrAy: Skip -> [" & M & "]"
        GoTo Nxt
    End If
    XX = GetAttr(A & M)
    If M = "." Then GoTo Nxt
    If M = ".." Then GoTo Nxt
    If XX And X Then
        Push O, M
    End If
Nxt:
    M = Dir
Wend
PthSubFdrAy = O
End Function

Function PthSubPthAy(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
PthSubPthAy = AyAddPfxSfx(PthSubFdrAy(A, Spec, Atr), A, "\")
End Function

Function TmpFb$(Optional Fdr$, Optional Fnn$)
TmpFb = TmpFfn(".accdb", Fdr, Fnn)
End Function

Function TmpFt$(Optional Fdr$, Optional Fnn$)
TmpFt = TmpFfn(".txt", Fdr, Fnn)
End Function

Function TmpFx$(Optional Fdr$, Optional Fnn$)
TmpFx = TmpFfn(".xlsx", Fdr, Fnn)
End Function
Function TmpFxa$(Optional Fdr$, Optional Fnn$)
TmpFxa = TmpFfn(".xlam", Fdr, Fnn)
End Function

Function TmpNm$()
Static X&
TmpNm = "T" & Format(Now(), "YYYYMMDD_HHMMSS") & "_" & X
X = X + 1
End Function

Function TmpPth$(Optional Fdr$)
Dim X$
   If Fdr <> "" Then
       X = Fdr & "\"
   End If
Dim O$
   O = TmpPthFix & X:   PthEns O
   O = O & TmpNm & "\": PthEns O
   PthEns O
TmpPth = O
End Function

Sub TmpPthBrw()
PthBrw TmpPth
End Sub

Function TmpPthFix$()
Static X$
If X = "" Then X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
TmpPthFix = X
End Function

Private Sub PthPushEntAyR(A)
'Debug.Print "PthPUshEntAyR:" & A
Dim P$(): P = PthSubPthAy(A)
If Sz(P) = 0 Then Exit Sub
If Sz(O) Mod 1000 = 0 Then Debug.Print "PthPushEntAyR: (Each 1000): " & A
PushAy O, P
PushAy O, PthFfnAy(A)
Dim PP
For Each PP In P
    PthPushEntAyR PP
Next
End Sub

Private Sub PthEntAy__Tst()
Dim A$(): A = PthEntAy("C:\users\user\documents\", IsRecursive:=True)
Debug.Print Sz(A)
Stop
AyDmp A
End Sub

Private Sub PthRmvEmpSubDir__Tst()
TmpPth
End Sub


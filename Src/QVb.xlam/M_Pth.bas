Attribute VB_Name = "M_Pth"
Option Explicit
Private O$() ' Used by PthEntAyR

Property Get PthEntAy(A, Optional FilSpec$ = "*.*", Optional Atr As FileAttribute, Optional IsRecursive As Boolean) As String()
If Not IsRecursive Then
    PthEntAy = AyAdd(PthSubPthAy(A), PthFfnAy(A, FilSpec, Atr))
    Exit Property
End If

Erase O
PushEntAyR A
PthEntAy = O
Erase O
End Property

Property Get PthFdr$(A)
PthFdr = TakAftRev(RmvLasChr(A), "\")
End Property

Property Get PthFfnAy(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
PthFfnAy = AyAddPfx(PthFnAy(A, Spec, Atr), A)
End Property

Property Get PthFnAy(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
Ass PthIsExist(A)
Dim O$()
Dim M$
M = Dir(A & Spec)
If Atr = 0 Then
    While M <> ""
        Stop
'       M_Ay.Push O, M
       M = Dir
    Wend
    PthFnAy = O
End If
While M <> ""
    If GetAttr(A & M) And Atr Then
        Push O, M
    End If
    M = Dir
Wend
PthFnAy = O
End Property

Property Get PthHasFil(A) As Boolean
If Not PthIsExist(A) Then Exit Property
PthHasFil = (Dir(A & "*.*") <> "")
End Property

Property Get PthHasPthSfx(A) As Boolean
PthHasPthSfx = LasChr(A) = "\"
End Property

Property Get PthHasSubDir(A) As Boolean
If Not PthIsExist(A) Then Exit Property
Dim P$: P = Dir(A & "*.*", vbDirectory)
PthHasSubDir = Dir <> ""
End Property

Property Get PthIsEmp(A) As Boolean
If PthHasFil(A) Then Exit Property
If PthHasSubDir(A) Then Exit Property
PthIsEmp = True
End Property

Property Get PthIsExist(A) As Boolean
PthIsExist = Fso.FolderExists(A)
End Property

Property Get PthSubFdrAy(A, Optional Spec$ = "*.*", Optional Atr As VbFileAttribute) As String()
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
        Stop
        'Push O, M
    End If
Nxt:
    M = Dir
Wend
PthSubFdrAy = O
End Property

Property Get PthSubPthAy(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
PthSubPthAy = AyAddPfxSfx(PthSubFdrAy(A, Spec, Atr), A, "\")
End Property

Sub PthBrw(A)
Shell "Explorer """ & A & """", vbMaximizedFocus
End Sub

Sub PthClrFil(A)
If Not PthIsExist(A) Then Exit Sub
Dim F
Dim Ay$(): Ay = PthFfnAy(A)
If AyIsEmp(Ay) Then Exit Sub
For Each F In Ay
   Kill F
Next
End Sub

Sub PthEns(A)
If PthIsExist(A) Then Exit Sub
MkDir A
End Sub

Sub PthRmvEmpSubDir(A)
Dim P$(): P = PthSubPthAy(A)
If AyIsEmp(P) Then Exit Sub
Dim I
For Each I In P
   PthRmvIfEmp A
Next
End Sub

Sub PthRmvIfEmp(A)
If Not PthIsExist(A) Then Exit Sub
If PthIsEmp(A) Then Exit Sub
RmDir A
End Sub

Sub ZZ__Tst()
ZZ_PthEntAy
ZZ_PthRmvEmpSubDir
End Sub

Private Sub PushEntAyR(A)
Stop
'Debug.Print "PthPUshEntAyR:" & A
'Dim P$(): P = Path(A).SubPthAy
'If Sz(P) = 0 Then Exit Sub
'If Sz(O) Mod 1000 = 0 Then Debug.Print "PthPushEntAyR: (Each 1000): " & A
'PushAy O, P
'PushAy O, PthFfnAy(A)
'Dim PP
'For Each PP In P
'    PthPushEntAyR PP
'Next
End Sub

Private Sub ZZ_PthEntAy()
Dim A$(): A = PthEntAy("C:\users\user\documents\", IsRecursive:=True)
Debug.Print Sz(A)
Stop
AyDmp A
End Sub

Private Sub ZZ_PthRmvEmpSubDir()
PthRmvEmpSubDir TmpPth
End Sub

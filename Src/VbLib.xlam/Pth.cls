VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A$
Private O$() ' Used by PthEntAyR
Friend Property Get Init(Pth) As Pth
A = Pth
Ass HasPthSfx
Set Init = Me
End Property

Sub Brw()
Shell "Explorer """ & TmpPth & """", vbMaximizedFocus
End Sub

Sub ClrFil()
If Not IsExist Then Exit Sub
Dim F
Dim Ay$(): Ay = FfnAy
If AyIsEmp(Ay) Then Exit Sub
For Each F In Ay
   Kill F
Next
End Sub

Sub Ens()
If IsExist Then Exit Sub
MkDir A
End Sub

Function EntAy(Optional FilSpec$ = "*.*", Optional Atr As FileAttribute, Optional IsRecursive As Boolean) As String()
If Not IsRecursive Then
    Stop
'    PthEntAy = AyAdd(PthSubPthAy(A), PthFfnAy(FilSpec, Atr))
    Exit Function
End If

Erase O
PushEntAyR A
EntAy = O
Erase O
End Function

Function Fdr$()
Stop
'Dim P$: P = RmvLasChr(A)
'PthFdr = TakAftRev(A, "\")
End Function

Function FfnAy(Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
Stop
'FfnAy = AyAddPfx(FnAy(Spec, Atr), A)
End Function

Function FnAy(Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
Ass IsExist
Dim O$()
Dim M$
M = Dir(A & Spec)
If Atr = 0 Then
    While M <> ""
        Stop
'       M_Ay.Push O, M
       M = Dir
    Wend
    FnAy = O
End If
While M <> ""
    If GetAttr(A & M) And Atr Then
'        M_Ay.Push O, M
    End If
    M = Dir
Wend
FnAy = O
End Function

Function HasFil() As Boolean
If Not IsExist Then Exit Function
HasFil = (Dir(A & "*.*") <> "")
End Function

Function HasPthSfx() As Boolean
HasPthSfx = LasChr(A) = "\"
End Function

Function HasSubDir() As Boolean
If Not IsExist Then Exit Function
Dim P$: P = Dir(A & "*.*", vbDirectory)
Dir
HasSubDir = Dir <> ""
End Function

Function IsEmp() As Boolean
If HasFil Then Exit Function
If HasSubDir Then Exit Function
IsEmp = True
End Function

Function IsExist() As Boolean
IsExist = Fso.FolderExists(A)
End Function

Sub RmvEmpSubDir()
Dim P$(): P = SubPthAy
If AyIsEmp(P) Then Exit Sub
Dim I
For Each I In P
   Vb.Pth(I).RmvIfEmp
Next
End Sub

Sub RmvIfEmp()
If Not IsExist Then Exit Sub
If IsEmp Then Exit Sub
RmDir A
End Sub

Function SubFdrAy(Optional Spec$ = "*.*", Optional Atr As VbFileAttribute) As String()
'PthSubFdrAy = ItrNy(Fso.GetFolder(A).SubFolders, Spec)
Ass IsExist
Ass HasPthSfx
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
SubFdrAy = O
End Function

Function SubPthAy(Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
Stop
'SubPthAy = AyAddPfxSfx(SubFdrAy(Spec, Atr), A, "\")
End Function

Private Sub PushEntAyR(A)
Stop
'Debug.Print "PthPUshEntAyR:" & A
'Dim P$(): P = Pth(A).SubPthAy
'If Sz(P) = 0 Then Exit Sub
'If Sz(O) Mod 1000 = 0 Then Debug.Print "PthPushEntAyR: (Each 1000): " & A
'PushAy O, P
'PushAy O, PthFfnAy(A)
'Dim PP
'For Each PP In P
'    PthPushEntAyR PP
'Next
End Sub

Private Sub ZZ_EntAy()
Dim A$(): A = Pth("C:\users\user\documents\").EntAy(IsRecursive:=True)
'Debug.Print Sz(A)
Stop
'AyDmp A
End Sub

Private Sub ZZ_RmvEmpSubDir()
TmpPth
End Sub

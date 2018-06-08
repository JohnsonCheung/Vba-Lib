VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A$()
Property Get U&()
U = UB(A)
End Property
Function HasMajPfx(MajPfx$) As Boolean
Dim Cnt%, J%
For J = 0 To U
    If HasPfx(A(J), MajPfx) Then Cnt = Cnt + 1
Next
HasMajPfx = Cnt > (Sz(A) \ 2)
End Function

Function LnxsAy(LinPfx$) As Lnx()
Stop
End Function
Property Get Lnxs() As Lnxs
Dim O As New Lnxs
Set Lnxs = O.Init(LnxAy)
End Property
Property Get LnxAy() As Lnx()
Dim O() As Lnx
Dim J%
For J = 0 To U
    PushObj O, Vb.Lnx(A(J), J)
Next
LnxAy = O
End Property

Property Get Dic(Optional JnSep$ = vbCrLf) As Dictionary
Const CSub$ = "Ly.Dic"
Dim O As New Dictionary
   If AyIsEmp(A) Then Set Dic = O: Exit Property
   Dim I
   For Each I In A
       If Trim(I) = "" Then GoTo Nxt
       If FstChr(I) = "#" Then GoTo Nxt
       With Brk(I, " ")
           If O.Exists(.S1) Then
               O(.S1) = O(.S1) & JnSep & .S2
           Else
               O.Add .S1, .S2
           End If
       End With
Nxt:
   Next
Set Dic = O
End Property

Property Get Init(Ly$()) As Ly
A = Ly
Set Init = Me
End Property
Property Get ToStr$()
If IsEmp Then
    ToStr = "Ly()"
Else
    ToStr = FmtQQ("Ly(|?|)", JnCrLf(A, WithIx:=True))
End If
End Property
Property Get N&()
N = Sz(A)
End Property
Property Get IsEmp() As Boolean
IsEmp = N = 0
End Property
Property Get T1AyRstAy() As SyPair
Dim J&, T1$(), Rst$(), L As Lin
For J = 0 To U
    Set L = Lin(A(J))
    With L.T1Rst
        Push T1, .T1
        Push Rst, .Rst
    End With
Next
With T1AyRstAy
    .Sy1 = T1
    .Sy2 = Rst
End With
End Property
Friend Sub Tst()
ZZ_T1AyRstAy
End Sub

Private Sub ZZ_T1AyRstAy()
Dim A$()
Push A, "lskdfj sldkfj sldfj sldkfj sldf j"
Push A, "lksj flskdj flsdjk fsldjkf"
Dim Act As SyPair
Act = Ly(A).T1AyRstAy
Stop
End Sub

Property Get EndTrim() As String()
If IsEmp Then Exit Property
If Not Lin(AyLasEle(A)).IsEmp Then EndTrim = A: Exit Property
Dim J%
For J = U To 0 Step -1
    If Not Lin(A(J)).IsEmp Then
        Dim O$()
        O = A
        ReDim Preserve O(J)
        EndTrim = O
        Exit Property
    End If
Next
End Property


Property Get LyRmv2Dash(Ly$()) As String()
If AyIsEmp(Ly) Then Exit Property
Dim O$(), I
For Each I In Ly
    Push O, Rmv2Dash(CStr(I))
Next
LyRmv2Dash = O
End Property


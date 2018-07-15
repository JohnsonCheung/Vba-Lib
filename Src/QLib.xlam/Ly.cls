VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private B_Ly$()

Property Get Dic(Optional JnSep$ = vbCrLf) As Dictionary
Const CSub$ = "Ly.Dic"
Dim O As New Dictionary
   If AyIsEmp(B_Ly) Then Set Dic = O: Exit Property
   Dim I
   For Each I In B_Ly
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
B_Ly = Ly
Set Init = Me
End Property

Property Get IsEmp() As Boolean
IsEmp = N = 0
End Property

Property Get LnxAy() As Lnx1()
Dim O() As Lnx1
Dim J%
For J = 0 To U
    PushObj O, Vb.Lnx(B_Ly(J), J)
Next
LnxAy = O
End Property

Property Get Lnxs() As Lnx1s
Dim O As New Lnx1s
Set Lnxs = O.Init(LnxAy)
End Property

Property Get N&()
N = Sz(B_Ly)
End Property

Property Get T1AyRstAy() As T1AyRstAy
Dim J&, T1$(), Rst$(), L As Lin
For J = 0 To U
    Set L = Lin(B_Ly(J))
    With L.T1Rst
        Push T1, .T1
        Push Rst, .Rst
    End With
Next
With T1AyRstAy
    .T1Ay = T1
    .RstAy = Rst
End With
End Property

Property Get ToStr$()
If IsEmp Then
    ToStr = "Ly()"
Else
    ToStr = FmtQQ("Ly(|?|)", JnCrLf(B_Ly, WithIx:=True))
End If
End Property

Property Get U&()
U = N - 1
End Property

Function LnxsAy(LinPfx$) As Lnx1()
Stop
End Function

Friend Sub Tst()
ZZ_T1AyRstAy
End Sub

Private Sub ZZ_T1AyRstAy()
Dim A$()
Push A, "lskdfj sldkfj sldfj sldkfj sldf j"
Push A, "lksj flskdj flsdjk fsldjkf"
Dim Act As T1AyRstAy
Act = Ly(A).T1AyRstAy
Stop
End Sub

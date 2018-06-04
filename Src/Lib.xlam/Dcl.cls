VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dcl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private A$()
Friend Property Get Init(Dcl$()) As Dcl
A = Dcl
Set Init = Me
End Property
Property Get N%()
N = Sz(A)
End Property
Private Property Get U%()
U = N - 1
End Property
Function EnmBdyLy(EnmNm$) As String()
Dim B%: B = EnmLx(EnmNm): If B = -1 Then Exit Function
Dim O$(), J%
For J = B To U
   Push O, A(J)
   If HasPfx(A(J), "End Enum") Then EnmBdyLy = O: Exit Function
Next
Stop
End Function

Function EnmLx%(EnmNm$)
Dim O%, L$
For O = 0 To U
If SrcLin_IsEmn(A(O)) Then
    L = A(O)
    L = SrcLin_RmvMdy(L)
    With Lin(L)
         L = .RmvT1
         If .T1 = EnmNm Then
            EnmLx = O: Exit Function
         End If
    End With
End If
Next
EnmLx = -1
End Function

Property Get EnmNy() As String()
If AyIsEmp(A) Then Exit Property
Dim I, O$()
For Each I In A
   PushNonEmp O, NewSrcLin(I).EnmNm
Next
EnmNy = O
End Property

Function HasTy(TyNm$) As Boolean
If AyIsEmp(A) Then Exit Function
Dim I
For Each I In A
   If HasPfx(I, "Type") Then If Lin(I).T2 = TyNm Then HasTy = True: Exit Function
Next
End Function
Private Sub TyLines__Tst()
Debug.Print Ide.Dcl(MdDclLy(CurMd)).TyLines("AA")
End Sub

Property Get NEnm%()
If AyIsEmp(A) Then Exit Property
Dim I, O%
For Each I In A
   If SrcLin_IsEmn(I) Then O = O + 1
Next
NEnm = O
End Property

Function TyFmIx&(TyNm$)
Dim J%, L$
For J = 0 To UB(A)
   If SrcLin_TyNm(A(J)) = TyNm Then TyFmIx = J: Exit Function
Next
TyFmIx = -1
End Function

Function TyFmTo(TyNm$) As FmTo
Dim FmI&: FmI = TyFmIx(TyNm)
Dim ToI&: ToI = TyToIx(FmI)
TyFmTo = NewFmTo(FmI, ToI)
End Function

Function TyIx%(TyNm)
Dim I%
   For I = 0 To UB(A)
       If SrcLin_TyNm(A(I)) = TyNm Then
           TyIx% = I
           Exit Function
       End If
   Next
TyIx = -1
End Function

Function TyLines$(TyNm$)
Stop
TyLines = JnCrLf(TyLy(TyNm))
End Function

Function TyLy(TyNm$) As String()
TyLy = AyWhFmTo(A, TyFmTo(TyNm))
End Function

Property Get TyNy(Optional TyNmPatn$ = ".") As String()
If AyIsEmp(A) Then Exit Property
Dim O$(), L, M$, R As Re
Set R = Re(TyNmPatn)
For Each L In A
   M = SrcLin_TyNm(L)
   If R.Tst(M) Then
       PushNonEmp O, M
   End If
Next
TyNy = O
End Property

Private Function TyToIx(FmI&)
If 0 > FmI Then TyToIx = -1: Exit Function
Dim O&
For O = FmI + 1 To UB(A)
   If HasPfx(A(O), "End Type") Then TyToIx = O: Exit Function
Next
TyToIx = -1
End Function


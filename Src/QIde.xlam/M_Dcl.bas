Attribute VB_Name = "M_Dcl"
Option Explicit

Function DclEnmLx%(A$(), EnmNm$)
Dim O%, L$
For O = 0 To U
If SrcLin_IsEmn(A(O)) Then
    L = A(O)
    L = SrcLin_RmvMdy(L)
    With Lin(L)
         L = .RmvT1
         If .T1 = EnmNm Then
            DclEnmLx = O: Exit Function
         End If
    End With
End If
Next
DclEnmLx = -1
End Function

Function DclHasTy(A$(), TyNm$) As Boolean
If AyIsEmp(A) Then Exit Function
Dim I
For Each I In A
   If HasPfx(I, "Type") Then If Lin(I).T2 = TyNm Then DclHasTy = True: Exit Function
Next
End Function

Function DclTyFmIx%(A$(), TyNm$)
Dim J%, L$
For J = 0 To UB(A)
   If SrcLin_TyNm(A(J)) = TyNm Then DclTyFmIx = J: Exit Function
Next
DclTyFmIx = -1
End Function

Function DclTyFmTo(A$(), TyNm$) As FmTo
Dim FmI&: FmI = DclTyFmIx(A, TyNm)
Dim ToI&: ToI = DclTyToIx(A, FmI)
TyFmTo = NewFmTo(FmI, ToI)
End Function

Function DclTyLines$(A$(), TyNm$)
DclTyLines = JnCrLf(DclTyLy(A, TyNm))
End Function

Function DclTyLy(A$(), TyNm$) As String()
DclTyLy = AyWhFmTo(A, DclTyFmTo(TyNm))
End Function

Function DclEnmBdyLy(A$(), EnmNm$) As String()
Dim B%: B = EnmLx(EnmNm): If B = -1 Then Exit Function
Dim O$(), J%
For J = B To U
   Push O, A(J)
   If HasPfx(A(J), "End Enum") Then EnmBdyLy = O: Exit Function
Next
Stop
End Function

Function DclEnmNy(A$()) As String()
If AyIsEmp(A) Then Exit Function
Dim I, O$()
For Each I In A
   PushNonEmp O, NewSrcLin(I).EnmNm
Next
DclEnmNy = O
End Function

Function DclNEnm%(A$())
If AyIsEmp(A) Then Exit Function
Dim I, O%
For Each I In A
   If SrcLin_IsEmn(I) Then O = O + 1
Next
NEnm = O
End Function

Function DclTyNy(A$(), Optional TyNmPatn$ = ".") As String()
If AyIsEmp(A) Then Exit Function
Dim O$(), L, M$, R As Re
Set R = Re(TyNmPatn)
For Each L In A
   M = SrcLin_TyNm(L)
   If R.Tst(M) Then
       PushNonEmp O, M
   End If
Next
TyNy = O
End Function

Private Function DclTyToIx%(A$(), FmLx%)
If 0 > FmI Then TyToIx = -1: Exit Function
Dim O&
For O = FmI + 1 To UB(A)
   If HasPfx(A(O), "End Type") Then TyToIx = O: Exit Function
Next
DclTyToIx = -1
End Function

Private Sub ZZ_DclTyLines()
Debug.Print DclTyLines(MdDclLy(CurMd), "AA")
End Sub

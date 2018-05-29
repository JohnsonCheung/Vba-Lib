Attribute VB_Name = "IdeEnsDclOpt"
Option Explicit
Private Enum eOptTy
   eExplicit = 1
   eCmpDb = 2
End Enum
Private A_Md As CodeModule
Private A_Ty As eOptTy

Sub EnsMdCmpDb(Optional Nm$)
EnsMd Nm, eCmpDb
End Sub

Sub EnsMdExplicit(Optional Nm$)
EnsMd Nm, eExplicit
End Sub

Sub EnsPjCmpDb(Optional Nm$)
EnsPj Nm, eCmpDb
End Sub

Sub EnsPjExplicit(Optional Nm$)
EnsPj Nm, eExplicit
End Sub

Private Sub Ens()
If HasLin Then
   Debug.Print FmtQQ("[?] already exists in Md-[?] ", OptLin, MdNm(A_Md))
   Exit Sub
End If
Debug.Print FmtQQ("[?] of Md-[?] is added ........................", OptLin, MdNm(A_Md))
Dim I%
A_Md.InsertLines Ix, OptLin
End Sub

Private Sub EnsMd(Nm$, T As eOptTy)
A_Ty = T
Set A_Md = DftMdByMdNm(Nm)
Ens
End Sub

Private Sub EnsPj(Nm$, T As eOptTy)
Dim Pj As VBProject
'   Set Pj = DftPjByPjNm(Nm)
Dim I
A_Ty = T
'For Each I In PjMbrAy(Pj)
   Set A_Md = I
   Ens
'Next
End Sub

Private Property Get HasLin() As Boolean
Dim Dcl$()
   Dcl = MdDcl(A_Md)
Dim A$
   A = OptLin
Dim J%
For J = 0 To UB(Dcl)
   If HasPfx(Dcl(J), A) Then HasLin = True: Exit Property
Next
End Property

Private Property Get Ix%()
Dim J%
For J% = 1 To A_Md.CountOfDeclarationLines
   Dim L$
   L = A_Md.Lines(J, 1)
   If Lin(L).IsEmp Then Ix = J
   If Not SrcLin_IsRmk(L) Then Ix = J
Next
Ix = J
End Property

Private Property Get OptLin$()
Select Case A_Ty
Case eCmpDb: OptLin = "Option Compare Database"
Case eExplicit: OptLin = "Option Explicit"
End Select
End Property


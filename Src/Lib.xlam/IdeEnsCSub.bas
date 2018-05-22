Attribute VB_Name = "IdeEnsCSub"
Option Explicit
Type Drec
   Fny() As String
   Dr() As Variant
End Type
Private Type CSubBrk
   NeedDlt As Boolean
   NeedIns As Boolean
   OldLno As Long
   OldCSub As String
   NewLno As Long
   NewCSub As String
   MdNm As String
   MthNm As String
End Type

Function CSubBrk_Str$(A As CSubBrk)
With A
   CSubBrk_Str = JnTab(Array(IIf(.OldCSub = .NewCSub, "*NoChg", "*Upd"), A.MdNm, AlignL(A.MthNm, 25, DoNotCut:=True), "CSub=[" & .NewCSub & "]"))
End With
End Function

Sub MdBrwCSubDrs(A As CodeModule)
DrsBrw MdCSubDrs(A)
End Sub

Sub MdEnsCSub(A As CodeModule)
Dim Mth, Ny$()
Ny = MdMthNy(A)
If AyIsEmp(Ny) Then Exit Sub
For Each Mth In Ny
   MdMth_EnsCSub A, CStr(Mth)
Next
End Sub

Function MdMth_CSubBrk(A As CodeModule, MthNm$) As CSubBrk
Const CSub$ = "MdMth_CSubBrk"
Dim MLy$()
Dim MLno&
   MLno = MdMth_Lno(A, MthNm)
   MLy = MdMth_BdyLy(A, MthNm)

Dim IsUsingCSub As Boolean '-> NewAt
   IsUsingCSub = False
   If HasSubStrAy(Join(MLy), Array("Er CSub,", "Debug.Print CSub", "(CSub,")) Then
       IsUsingCSub = True
   End If

Dim OldCSubIx%
   Dim J%
   OldCSubIx = -1
   For J = 0 To UB(MLy)
       If HasPfx(MLy(J), "Const CSub") Then
           OldCSubIx = J
       End If
   Next

Dim OOldLno&
   OOldLno = IIf( _
       OldCSubIx >= 0, _
       MLno + OldCSubIx, _
       0)

Dim OOldCSub$
   If OldCSubIx >= 0 Then
       OOldCSub = MLy(OldCSubIx)
   Else
       OOldCSub = ""
   End If

Dim ONewLno&
   If IsUsingCSub Then
       Dim Fnd As Boolean
       For J = 0 To UB(MLy)
           If LasChr(MLy(J)) <> "_" Then
               Fnd = True
               ONewLno = MLno + J + 1
               Exit For
           End If
       Next
       If Not Fnd Then Er CSub, "{MthLy} has all lines with _ as sfx with is impossible", MLy
   Else
       ONewLno = 0
   End If

Dim ONewCSub$

   If ONewLno > 0 Then
       ONewCSub = FmtQQ("Const CSub$ = ""?""", A)
   Else
       ONewCSub = ""
   End If

Dim O As CSubBrk
   Dim HasOldCSub As Boolean
   Dim HasNewCSub As Boolean
   Dim IsDiff As Boolean
       HasOldCSub = OOldCSub <> ""
       HasNewCSub = ONewCSub <> ""
       IsDiff = OOldCSub <> ONewCSub
   With O
       .NeedDlt = IsDiff And HasOldCSub
       .NeedIns = IsDiff And HasNewCSub
       .NewCSub = ONewCSub
       .NewLno = ONewLno
       .OldCSub = OOldCSub
       .OldLno = OOldLno
       .MdNm = MdNm(A)
       .MthNm = MthNm
   End With
MdMth_CSubBrk = O
End Function

Sub MdMth_DmpCSubBrk(A As CodeModule, MthNm$)
Const CSub$ = "DmpCSubBrk"
Debug.Print CSubBrk_Str(MdMth_CSubBrk(A, MthNm))
End Sub

Sub MdMth_EnsCSub(A As CodeModule, MthNm$)
Const CSub$ = "MdMthEnsCSub"
Dim B As CSubBrk
    B = MdMth_CSubBrk(A, MthNm)
With B
   If .NeedDlt Then
       A.DeleteLines .OldLno         '<==
   End If
   If .NeedIns Then
       A.InsertLines .NewLno, .NewCSub
   End If
End With
Debug.Print CSubBrk_Str(B)
End Sub

Function PjCSubDt(A As VBProject) As Dt
Dim I, Md As CodeModule
Dim Dry()
For Each I In Pj(A).MdAy
   Set Md = I
   PushAy Dry, MdCSubDry(Md)
Next
PjCSubDt = NewDt("Pj-CSub", FnyOf_CSub, Dry)
End Function

Sub PjEnsCSub(A As VBProject)
Dim I, Md As CodeModule
For Each I In Pj(A).MdAy
   Set Md = I
   MdEnsCSub Md
Next
End Sub

Private Sub CSubBrk_Dmp(A As CSubBrk)
DrecDmp CSubBrk_Drec(A)
End Sub

Private Function CSubBrk_Dr(A As CSubBrk) As Variant()
With A
   CSubBrk_Dr = Array(.MdNm, .MthNm, .NeedDlt, .NeedIns, .NewLno, .NewCSub, .OldLno, .OldCSub)
End With
End Function

Private Function CSubBrk_Drec(A As CSubBrk) As Drec
With A
   CSubBrk_Drec.Dr = CSubBrk_Dr(A)
   CSubBrk_Drec.Fny = FnyOf_CSub
End With
End Function

Private Function FnyOf_CSub() As String()
FnyOf_CSub = LvsSy("MdNm MthNm NeedDlt NeedIns NewLno NewCSub OldLno OldCSub")
End Function

Private Function MdCSubDrs(A As CodeModule) As Drs
MdCSubDrs.Fny = FnyOf_CSub
MdCSubDrs.Dry = MdCSubDry(A)
End Function

Private Function MdCSubDry(A As CodeModule) As Variant()
Dim Mth, Dry()
Dim Ny$(): Ny = MdMthNy(A)
If AyIsEmp(Ny) Then Exit Function
For Each Mth In Ny
   Push Dry, CSubBrk_Dr(MdMth_CSubBrk(A, CStr(Mth)))
Next
MdCSubDry = Dry
End Function

Private Sub CSubBrk_Dmp__Tst()
End Sub

Private Sub MdMth_CSubBrk__Tst()
CSubBrk_Dmp MdMth_CSubBrk(CurMd, "MdMth_CSubBrk_Tst")
Stop
End Sub

Attribute VB_Name = "M_Mth"
Option Explicit
Type MthBrk
    MthNm As String
    Ty As String    ' Sub | Function | Property Get | Property Set | Property Let (Ty here means MthTy)
    Mdy As String
End Type
Type MthNmBrk
    MdTy As String
    MdNm As String
    MthNm As String
    Ty As String
    Mdy As String
End Type
Type PrmTy
    TyChr As String
    TyAsNm As String
    IsAy As Boolean
End Type
Type MthPrm
    Nm As String
    IsOpt As Boolean
    IsPrmAy As Boolean
    Ty As PrmTy
    DftVal As String
End Type
Type MthSig
    HasRetVal As Boolean
    PrmAy() As MthPrm
    RetTy As PrmTy
End Type







Function MthBdyLines$(A As CodeModule, MthNm$)
MthBdyLines = SrcMth_BdyLines(MdBdyLy(A), MthNm)
End Function


Function MthLinArgStr$(MthLin$)
MthLinArgStr = TakBetBkt(MthLin)
End Function

Function MthLinHasRetVal(MthLin$ _
) As Boolean
Const CSub$ = "MthLinHasRetVal"
Dim A As MthBrk
    A = SrcLin_MthBrk(MthLin)
Select Case A.Ty
Case "Function", "Get": MthLinHasRetVal = True
Case "": Er CSub, "Give {MthLin} is not MthLin", MthLin
End Select
End Function

Function MthLinPrmAy(MthLin$) As MthPrm()
Dim ArgStr$
    ArgStr = TakBetBkt(MthLin, "()")
Dim P$()
    P = SplitComma(ArgStr)
Dim O() As MthPrm
    Dim U%: U = UB(P)
    ReDim O(U)
    Dim J%
    For J = 0 To U
        O(J) = NewMthPrm(P(J))
    Next
MthLinPrmAy = O
End Function

Function MthLinRetTy(MthLin$) As PrmTy
If Not HasSubStr(MthLin, "(") Then Exit Function
If Not HasSubStr(MthLin, ")") Then Exit Function
Dim TC$: TC = LasChr(TakBefBkt(MthLin))
With MthLinRetTy
    If IsTyChr(TC) Then .TyChr = TC: Exit Function
    Dim Aft$: Aft = TakAftBkt(MthLin)
        If Aft = "" Then Exit Function
        If Not HasPfx(Aft, " As ") Then Stop
        Aft = RmvPfx(Aft, " As ")
        If HasSfx(Aft, "()") Then
            .IsAy = True
            Aft = RmvSfx(Aft, "()")
        End If
        .TyAsNm = Aft
        Exit Function
End With
End Function

Function MthLno&(A As CodeModule, MthNm$)
MthLno = 1 + SrcMth_Lx(MdSrc(A), MthNm)
End Function

Function MthLnoAy(A As CodeModule, MthNm$) As Integer()
MthLnoAy = AyIncNForEachEle(SrcMth_LxAy(MdSrc(A), MthNm), 1)
End Function

Function MthMov(A As CodeModule, MthNm$, TarMd As CodeModule)
Ass Not IsNothing(A)
Ass Not IsNothing(TarMd)

Dim Bdy$: Bdy = MthBdyLines(A, MthNm)
If Bdy = "" Then Exit Function
TarMd.AddFromString Bdy
'MthRmv A, MthNm
End Function

Function MthPrm1(MthPrmStr$) As MthPrm
Stop '
'Const CSub$ = "MthPrm1"
'Dim A As Parse: NewParse (MthPrmStr)
'Dim TyChr$
'With MthPrm1
'    A = ParseKwOptional(A): .IsOpt = A.IsOk
'    A = ParseKwPrmAy(A):    .IsPrmAy = A.IsOk
'    A = ParseNm(A):       .Nm = ParseRet(A): If Not A.IsOk Then Er CSub, "A [Nm] is expected in {MthPrmStr}", MthPrmStr
'    A = ParseKwTyChr(A):    .Ty.TyChr = ParseRet(A)
'    A = ParseKwOptBktPair(A): .Ty.IsAy = ParseRet(A) = "()"
'End With
End Function

Function MthPrmSz&(A() As MthPrm)
On Error Resume Next
MthPrmSz = UBound(A) + 1
End Function

Function MthPrmUB&(A() As MthPrm)
MthPrmUB = MthPrmSz(A) - 1
End Function

Function MthSig(MthLin$) As MthSig
Dim O As MthSig
With O
    .HasRetVal = MthLinHasRetVal(MthLin)
    .PrmAy = MthLinPrmAy(MthLin)
    .RetTy = MthLinRetTy(MthLin)
End With
MthSig = O
End Function





Sub MthPrmPush(O() As MthPrm, I As MthPrm)
Dim N&: N = MthPrmSz(O)
ReDim Preserve O(N)
O(N) = I
End Sub

Sub MthSetMdy(A As CodeModule, MthNm$, Mdy$)
Ass KwIsMdy(Mdy)
Dim I&
    I = MthLno(A, MthNm)
Dim L$
    L = MdLin(A, I)
Dim Old$
    Old = SrcLin_Mdy(L)
If Mdy = Old Then Exit Sub
Dim NewL$
    Dim B$
    If Mdy <> "" Then
        B = Mdy & " "
    Else
        B = Mdy
    End If
    NewL = B & L
With A
    .DeleteLines I, 1
    .InsertLines I, NewL
End With
End Sub

Sub MthSetPrv(A As CodeModule, MthNm$)
MthSetMdy A, MthNm, "Private"
End Sub

Sub MthSetPub(A As CodeModule, MthNm$)
MthSetMdy A, MthNm, ""
End Sub


Private Sub MthBdyLines__Tst()
Debug.Print Len(MthBdyLines(CurMd, "MthLines"))
Debug.Print MthBdyLines(CurMd, "MthLines")
End Sub


Private Sub MthLinRetTy__Tst()
Dim MthLin$
Dim A As PrmTy:
MthLin = "Function MthPrm(MthPrmStr$) As MthPrm"
A = MthLinRetTy(MthLin)
Ass A.TyAsNm = "MthPrm"
Ass A.IsAy = False
Ass A.TyChr = ""

MthLin = "Function MthPrm(MthPrmStr$) As MthPrm()"
A = MthLinRetTy(MthLin)
Ass A.TyAsNm = "MthPrm"
Ass A.IsAy = True
Ass A.TyChr = ""

MthLin = "Function MthPrm$(MthPrmStr$)"
A = MthLinRetTy(MthLin)
Ass A.TyAsNm = ""
Ass A.IsAy = False
Ass A.TyChr = "$"

MthLin = "Function MthPrm(MthPrmStr$)"
A = MthLinRetTy(MthLin)
Ass A.TyAsNm = ""
Ass A.IsAy = False
Ass A.TyChr = ""
End Sub

Sub MthMov__Tst()
'MthMov Md("Mth_"), "XX", Md("A_")
End Sub
Function MthBdyLy(A As CodeModule, MthNm$) As String()
MthBdyLy = SrcMth_BdyLy(MdSrc(A), MthNm)
End Function
Function MthLnoCntAy(A As CodeModule, MthNm$) As LnoCnt()
MthLnoCntAy = SrcMth_LnoCntAy(MdSrc(A), MthNm)
End Function
Sub MthGo(A As CodeModule, MthNm$)
MdGoRRCC A, MthRRCC(A, MthNm)
End Sub
Sub MthRmv(A As CodeModule, MthNm$)
Dim M() As LnoCnt: M = MthLnoCntAy(A, MthNm)
If Sz(M) = 0 Then
    Debug.Print FmtQQ("Fun[?] in Md[?] not found, cannot Rmv", MthNm, MdNm(A))
Else
    Debug.Print FmtQQ("Fun[?] in Md[?] is removed", MthNm, MdNm(A))
End If
MdRmvLnoCntAy A, M
End Sub
Sub MthLnoCntAy__Tst()
Stop '
'Dim A() As LnoCnt: A = MthLnoCntAy(Md("Md_"), "XX")
'Dim J%
'For J = 0 To LnoCnt_UB(A)
'    LnoCnt_Dmp A(J)
'Next
End Sub
Function MthRRCC(A As CodeModule, MthNm$) As RRCC
Stop '
'MthRRCC = SrcMth_RRCC(MdSrc(A), MthNm)
End Function

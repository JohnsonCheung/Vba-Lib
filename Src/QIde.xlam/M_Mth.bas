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

Function CurTarMd() As CodeModule
With CurVbe
   If .CodePanes.Count <> 2 Then Exit Function
   Dim M1 As CodeModule: Set M1 = .CodePanes(1).CodeModule
   Dim M2 As CodeModule: Set M2 = .CodePanes(2).CodeModule
   Dim M As CodeModule: Set M = CurMd
   Dim IsM1Tar As Boolean: IsM1Tar = M1 <> M And M2 = M
   Dim IsM2Tar As Boolean: IsM2Tar = M2 <> M And M1 = M
   If Not (IsM1Tar Xor IsM2Tar) Then Stop
   If IsM1Tar Then Set CurTarMd = M1: Exit Function
   If IsM2Tar Then Set CurTarMd = M2: Exit Function
End With
End Function

Function IsOnlyTwoCdPne() As Boolean
IsOnlyTwoCdPne = CurVbe.CodePanes.Count = 2
End Function

Sub MovAllMth()

End Sub

Sub MovMth()

End Sub

Sub CurTarMd__Tst()
Debug.Print MdNm(CurTarMd)
End Sub

Function CurMthBdyLines$()
CurMthBdyLines = MdMth_BdyLines(CurMd, CurMthNm$)
End Function

Function CurMthNm$()
CurMthNm = MdCurMthNm(CurMd)
End Function

Function IsTstMthNm(MthNm$) As Boolean
IsTstMthNm = HasSfx(MthNm, "__Tst")
End Function

Function MdMthDic(A As CodeModule) As Dictionary
Set MdMthDic = SrcDic(MdSrc(A))
End Function

Function MdMth_BdyLines$(A As CodeModule, MthNm$)
MdMth_BdyLines = SrcMth_BdyLines(MdBdyLy(A), MthNm)
End Function

Function MdMth_Lno&(A As CodeModule, MthNm$)
MdMth_Lno = 1 + SrcMth_Lx(MdSrc(A), MthNm)
End Function

Function MdMth_LnoAy(A As CodeModule, MthNm$) As Integer()
MdMth_LnoAy = AyIncNForEachEle(SrcMth_LxAy(MdSrc(A), MthNm), 1)
End Function

Function MdMth_Mov(A As CodeModule, MthNm$, TarMd As CodeModule)
Ass Not IsNothing(A)
Ass Not IsNothing(TarMd)

Dim Bdy$: Bdy = MdMth_BdyLines(A, MthNm)
If Bdy = "" Then Exit Function
TarMd.AddFromString Bdy
'MdMth_Rmv A, MthNm
End Function

Sub MdMth_SetMdy(A As CodeModule, MthNm$, Mdy$)
Ass KwIsMdy(Mdy)
Dim I&
    I = MdMth_Lno(A, MthNm)
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

Sub MdMth_SetPrv(A As CodeModule, MthNm$)
MdMth_SetMdy A, MthNm, "Private"
End Sub

Sub MdMth_SetPub(A As CodeModule, MthNm$)
MdMth_SetMdy A, MthNm, ""
End Sub

Function MthBrk_Str$(A As MthBrk)
Dim O$()
PushNonEmp O, A.Mdy
PushNonEmp O, A.Ty
PushNonEmp O, A.MthNm
MthBrk_Str = JnSpc(O)
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

Function MthPrm1(MthPrmStr$) As MthPrm
Const CSub$ = "MthPrm1"
Dim A As Parse: NewParse (MthPrmStr)
Dim TyChr$
With MthPrm1
    A = ParseKwOptional(A): .IsOpt = A.IsOk
    A = ParseKwPrmAy(A):    .IsPrmAy = A.IsOk
    A = ParseNm(A):       .Nm = ParseRet(A): If Not A.IsOk Then Er CSub, "A [Nm] is expected in {MthPrmStr}", MthPrmStr
    A = ParseKwTyChr(A):    .Ty.TyChr = ParseRet(A)
    A = ParseKwOptBktPair(A): .Ty.IsAy = ParseRet(A) = "()"
End With
End Function

Sub MthPrmPush(O() As MthPrm, I As MthPrm)
Dim N&: N = MthPrmSz(O)
ReDim Preserve O(N)
O(N) = I
End Sub

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

Function NewMthPrm(MthPrmStr$) As MthPrm
Stop
'Dim L$: L = MthPrmStr
'Dim TyChr$
'With MthPrm
'    .IsOpt = ParseHasPfxSpc(L, "Optional")
'    .IsPrmAy = ParseHasPfxSpc(L, "ParamArray")
'    .Nm = ParseNm(L)
'    .Ty.TyChr = ParseOneOfChr(L, "!@#$%^&")
'    .Ty.IsAy = ParseHasPfx(L, "()")
'End With
End Function

Function PrmAyNy(A() As MthPrm) As String()
Dim J%, O$()
For J = 0 To MthPrmUB(A)
    Push O, A(J).Nm
Next
PrmAyNy = O
End Function

Function PrmTyAsTyNm$(A As PrmTy)
With A
    If .TyChr <> "" Then PrmTyAsTyNm = TyChrAsTyStr(.TyChr): Exit Function
    If .TyAsNm = "" Then
        PrmTyAsTyNm = "Variant"
    Else
        PrmTyAsTyNm = .TyAsNm
    End If
End With
End Function

Function PrmTyShtNm$(RetTy As PrmTy)
Dim Ay$
Dim O$
    With RetTy
        If .IsAy Then Ay = "Ay"
        Select Case .TyChr
        Case "!": O = "Sng"
        Case "@": O = "Cur"
        Case "#": O = "Dbl"
        Case "$": O = "Str"
        Case "%": O = "Int"
        Case "^": O = "LngLng"
        Case "&": O = "Lng"
        End Select
        If O = "" Then
            O = .TyAsNm
        End If
        If O = "" Then
            O = "Var"
        End If
    End With
    Select Case O
    Case "String": O = "Str"
    Case "Integer": O = "Int"
    Case "Long": O = "Lng"
    Case "Currency": O = "Cur"
    Case "Single": O = "Sng"
    Case "Double": O = "Dbl"
    Case "LongLong": O = "Lng"
    End Select
    O = O & Ay
    If O = "StrAy" Then O = "Sy"
PrmTyShtNm = O
End Function

Private Sub MdMth_BdyLines__Tst()
Debug.Print Len(MdMth_BdyLines(CurMd, "MdMth_Lines"))
Debug.Print MdMth_BdyLines(CurMd, "MdMth_Lines")
End Sub

Sub MdMth_Mov__Tst()
'MdMth_Mov Md("Mth_"), "XX", Md("A_")
End Sub

Sub MthDrs_SortingKy__Tst()
'AyDmp MthDrs_SortingKy(SrcMthDrs(MdSrc(Md("Mth_"))))
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


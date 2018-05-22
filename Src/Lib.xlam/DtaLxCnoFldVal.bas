Attribute VB_Name = "DtaLxCnoFldVal"
Option Explicit
Type VF
    V As String
    FldLvs As String
End Type
Type LFV
    Lx As Integer
    F As String
    V As String
End Type
Type LVF
    Lx As Integer
    V As String
    FldLvs As String
End Type
Type LCFV
    Lx As Integer
    Cno As Integer
    F As String
    V As String
End Type
Type LCFVRslt
    LCFVAy() As LCFV
    ErLy() As String
End Type
Const M_Should_Lng$ = "Lx(?) Fld(?) should have val(?) be a long number"
Const M_Should_Num$ = "Lx(?) Fld(?) should have val(?) be a number"
Const M_Should_Bet$ = "Lx(?) Fld(?) should have val(?) be between (?) and (?)"
Const M_Dup$ = _
                      "Lx(?) Fld(?) is found duplicated in Lx(?).  This item is ignored"
Sub LCFV_Push(O() As LCFV, A As LCFV)
Dim N%: N = LCFV_Sz(O)
ReDim Preserve O(N)
O(N) = A
End Sub
Sub LCFV_PushAy(O() As LCFV, A() As LCFV)
Dim J%
For J = 0 To LCFV_UB(A)
    LCFV_Push O, A(J)
Next
End Sub

Function LVF_Lin$(A As LVF)
With A
    LVF_Lin = FmtQQ("? ? ?", .Lx, .V, .FldLvs)
End With
End Function

Function LVFAy_Ly(A() As LVF) As String()
Dim O$(), J%
For J = 0 To LVF_UB(A)
    Push O, LVF_Lin(A(J))
Next
End Function

Sub LCFVRslt_Dmp(A As LCFVRslt)
'DsDmp LCFVRslt_Ds(A)
End Sub

Function LCFVRslt_Dt(A As LCFVRslt) As Dt
With LCFVRslt_Dt
    .Dry = LCFVRslt_Dry(A)
    .Fny = LvsSy("Lx Fld Cno Val")
    .DtNm = "LCFVRslt"
End With
End Function

Function LCFV_Sz%(A() As LCFV)
On Error Resume Next
LCFV_Sz = UBound(A) + 1
End Function

Function LCFV_UB%(A() As LCFV)
LCFV_UB = LCFV_Sz(A) - 1
End Function

Function LCFV_Has(A() As LCFV, M As LCFV) As Boolean
Dim J%, F$, V$
For J = 0 To LCFV_UB(A)
    If A(J).F = F Then
        If A(J).V = V Then
            LCFV_Has = True
            Exit Function
        End If
    End If
Next
End Function

Function LCFV_IsEmp(A() As LCFV) As Boolean
LCFV_IsEmp = LCFV_Sz(A) = 0
End Function

Function LCFV_Intersect(A() As LCFV, B() As LCFV) As LCFV()
If LCFV_IsEmp(A) Then Exit Function
If LCFV_IsEmp(A) Then Exit Function
Dim O() As LCFV
Dim J%
For J = 0 To LCFV_UB(A)
    If LCFV_Has(B, A(J)) Then LCFV_Push O, A(J)
Next
LCFV_Intersect = O
End Function

Function LCFV_Minus(A() As LCFV, B() As LCFV) As LCFV()

End Function

Function LCFVRslt_Dry(A As LCFVRslt) As Variant()
Dim O(), J%
For J = 0 To LCFV_UB(A.LCFVAy)
    Stop
Next
End Function

Function ErLy_Dt(ErLy$()) As Dt
ErLy_Dt = AyDt(ErLy, "Er", "Oup ErLy")
End Function

Function LCFVAy_CnoAy(A() As LCFV) As Integer()
Dim O%(), J%
For J = 0 To LCFV_UB(A)
    Push O, A(J).Cno
Next
LCFVAy_CnoAy = O
End Function

Function LCFVAy_FldValLy(A As LCFVRslt, T1$) As String()

End Function

Function LCFVAy_Fny(A() As LCFV) As String()
Dim O$(), J%
For J = 0 To LCFV_UB(A)
    With A(J)
        PushNoDup O, .F
    End With
Next
LCFVAy_Fny = O
End Function
Function LCFVAy_ValAy(A() As LCFV, OAy)
Erase OAy
Dim J%, O%()
For J = 0 To LCFV_UB(A)
    Push OAy, A(J).V
Next
LCFVAy_ValAy = OAy
End Function
Function LCFVAy_FldLvs$(A() As LCFV, V$)
Dim O$(), J%
For J = 0 To LCFV_UB(A)
    With A(J)
        If .V = V Then
            Push O, A(J).F
        End If
    End With
Next
If Sz(O) = 0 Then Stop
LCFVAy_FldLvs = JnSpc(O)
End Function

Function LCFVAy_VFAy(A() As LCFV) As VF()
Dim V$(): V = LCFVAy_ValAy(A, V)
Dim V1$(): V1 = AyUniq(V)
Dim O() As VF, FldLvs$, J%
For J = 0 To UB(V1)
    FldLvs = LCFVAy_FldLvs(A, V1(J))
    VF_PushVF O, V1(J), FldLvs
Next
LCFVAy_VFAy = O
End Function
Private Function VF(V$, FldLvs$) As VF
Dim O As VF
With O
    .FldLvs = FldLvs
    .V = V
End With
VF = O
End Function
Sub VF_PushVF(O() As VF, V$, FldLvs$)
VF_Push O, VF(V, FldLvs)
End Sub
Function VF_UB%(O() As VF)
VF_UB = VF_Sz(O) - 1
End Function

Function VF_Sz%(A() As VF)
On Error Resume Next
VF_Sz = UBound(A) + 1
End Function

Private Sub VF_Push(O() As VF, A As VF)
Dim N%: N = VF_Sz(O)
ReDim Preserve O(N)
O(N) = A
End Sub
Function LCFVAy_ValFldLy(A() As LCFV, T1$) As String()
Dim J%, O$(), VFAy() As VF
VFAy = LCFVAy_VFAy(A)
For J = 0 To VF_UB(VFAy)
    With VFAy(J)
        Push O, T1 & " " & .V & " " & .FldLvs
    End With
Next
LCFVAy_ValFldLy = O
End Function

Sub LCFVRslt_IODmp(A As LCFVRslt)
AyDmp LCFVRslt_IOLy(A)
End Sub

Function LCFVRslt_New(Ay() As LCFV) As LCFVRslt
LCFVRslt_New.LCFVAy = Ay
End Function

Function LCFVRslt_Add(A As LCFVRslt, B As LCFVRslt) As LCFVRslt
Dim O As LCFVRslt: O = A
With O
    LCFV_PushAy .LCFVAy, B.LCFVAy
    PushAy .ErLy, B.ErLy
End With
LCFVRslt_Add = O
End Function

Function LCFVRslt_IOLy(A As LCFVRslt) As String()
LCFVRslt_IOLy = DtLy(LCFVRslt_Dt(A))
End Function

Function LCFVRslt_RsltLy(A As LCFVRslt, InpDta() As LFV, InpFny$()) As String()
Dim O$()
Push O, "*Inp-Fny ========================================="
PushAy O, InpFny
PushAy O, DtLy(LFVAy_Dt(InpDta))
PushAy O, DtLy(LCFVRslt_Dt(A))
LCFVRslt_RsltLy = O
End Function

Function LCFVRslt_VdtIsNum(A As LCFVRslt) As LCFVRslt
Dim OAy() As LCFV
Dim OErLy$(), J%, Msg$
For J = 0 To LCFV_UB(A.LCFVAy)
    With A.LCFVAy(J)
        If IsNum(.V) Then
            LCFV_Push OAy, A.LCFVAy(J)
        Else
            Msg = FmtQQ(M_Should_Num, .Lx, .F, .V)
            Push OErLy, Msg
        End If
    End With
Next
Dim B As LCFVRslt
B.LCFVAy = OAy
B.ErLy = OErLy
LCFVRslt_VdtIsNum = LCFVRslt_Add(A, B)
End Function

Function LCFVRslt_VdtValBet(A As LCFVRslt, FmNum&, ToNum&) As LCFVRslt
Dim V&, J%, Msg$
Dim OErLy$(), OAy() As LCFV
For J = 0 To LCFV_UB(A.LCFVAy)
    With A.LCFVAy(J)
        V = Val(.V)
        If FmNum > V Or V > ToNum Then
            Msg = FmtQQ(M_Should_Bet, .Lx, .F, .V, FmNum, ToNum)
            Push OErLy, Msg
        Else
            LCFV_Push OAy, A.LCFVAy(J)
        End If
    End With
Next
Dim B As LCFVRslt
B.LCFVAy = OAy
B.ErLy = OErLy
LCFVRslt_VdtValBet = LCFVRslt_Add(A, B)
End Function

Function LCFVRslt_VdtVarIsLng(A As LCFVRslt) As LCFVRslt
Dim W%, J%, Msg$, OErLy$()
Dim OAy() As LCFV
For J = 0 To LCFV_UB(A.LCFVAy)
    With VarLngOpt(A.LCFVAy(J).V)
        If .Som Then
            LCFV_Push OAy, A.LCFVAy(J)
        Else
            With A.LCFVAy(J)
                Msg = FmtQQ(M_Should_Lng, .Lx, .F, .V)
            End With
            Push OErLy, Msg
        End If
    End With
Next
Dim B As LCFVRslt
B.ErLy = OErLy
B.LCFVAy = OAy
LCFVRslt_VdtVarIsLng = LCFVRslt_Add(A, B)
End Function

Private Function LFVAy_Dry(A() As LFV) As Variant()
Dim O(), J%
For J = 0 To LFV_UB(A)
    With A(J)
        Push O, Array(.Lx, .F, .V)
    End With
Next
LFVAy_Dry = O
End Function

Private Function LFVAy_Dt(A() As LFV) As Dt
With LFVAy_Dt
    .Dry = LFVAy_Dry(A)
    .Fny = LvsSy("Lx Fld Val")
    .DtNm = "Lx Fld Val"
End With
End Function

Function LFVAy_LCFVRslt(A() As LFV, Fny$()) As LCFVRslt
Dim A1 As LCFVRslt: A1 = LFVAy_VdtErFld(A, Fny)
LFVAy_LCFVRslt = LCFVRslt_VdtDupFld(A1)
End Function

Sub LFV_Push(O() As LFV, A As LFV)
Dim N%: N = LFV_Sz(O)
ReDim Preserve O(N)
O(N) = A
End Sub

Sub LFV_Push3(O() As LFV, Lx%, F$, V$)
Dim M As LFV
With M
    .F = F
    .Lx = Lx
    .V = V
End With
LFV_Push O, M
End Sub

Private Function LFV_Sz%(A() As LFV)
On Error Resume Next
LFV_Sz = UBound(A) + 1
End Function

Private Function LFV_UB%(A() As LFV)
LFV_UB = LFV_Sz(A) - 1
End Function

Sub LVFAy_Dmp(A() As LVF)
DrsDmp LVFAy_Drs(A)
End Sub

Function LVFAy_Drs(A() As LVF) As Drs
With LVFAy_Drs
    .Fny = LvsSy("Lx Val FldLvs")
    .Dry = LVFAy_Dry(A)
End With
End Function

Private Function LVFAy_Dry(A() As LVF) As Variant()
Dim O(), J%
For J = 0 To LVF_UB(A)
    With A(J)
        Push O, Array(.Lx, .V, .FldLvs)
    End With
Next
LVFAy_Dry = O
End Function

Function LVFAy_FldLvsAy(A() As LVF)
Dim O$(), J%
For J = 0 To LVF_UB(A)
    Push O, A(J).FldLvs
Next
LVFAy_FldLvsAy = O
End Function

Function LVFAy_Fny(A() As LVF) As String()
Dim O$(), J%
For J = 0 To LVF_UB(A)
    PushAy O, LvsSy(A(J).FldLvs)
Next
LVFAy_Fny = O
End Function

Function LVFAy_LCFVRslt(A() As LVF, Fny$()) As LCFVRslt
Dim A1() As LFV: A1 = LVFAy_LFVAy(A)
LVFAy_LCFVRslt = LFVAy_LCFVRslt(A1, Fny)
End Function

Sub LVF_Push(O() As LVF, A As LVF)
Dim N%: N = LVF_Sz(O)
ReDim Preserve O(N)
O(N) = A
End Sub

Sub LVF_Push3(O() As LVF, Lx%, V$, FldLvs$)
Dim M As LVF
With M
    .FldLvs = FldLvs
    .Lx = Lx
    .V = V
End With
LVF_Push O, M
End Sub

Function LCFVAy_WhByCnoAy(A() As LCFV, CnoAy%()) As LCFV()
Dim J%, O() As LCFV
For J = 0 To LCFV_UB(A)
    If AyHas(CnoAy, A(J).Cno) Then
        LCFV_Push O, A(J)
    End If
Next
LCFVAy_WhByCnoAy = O
End Function

Function LCFVRsltItm_Dry(A As LCFVRslt, Itm$) As Variant()
Dim O(), J%
For J = 0 To LCFV_UB(A.LCFVAy)
    With A.LCFVAy(J)
        Push O, Array(Itm, .Lx, .Cno, .F, .V)
    End With
Next
LCFVRsltItm_Dry = O
End Function

Function LVF_Sz%(A() As LVF)
On Error Resume Next
LVF_Sz = UBound(A) + 1
End Function

Function LVF_UB%(A() As LVF)
LVF_UB = LVF_Sz(A) - 1
End Function

Private Function LCFVAyIx_DupLxOpt(A() As LCFV, Ix%) As IntOpt
Dim I%, F$
F = A(Ix).F
For I = Ix + 1 To LCFV_UB(A)
    If F = A(I).F Then LCFVAyIx_DupLxOpt = SomInt(A(I).Lx): Exit Function
Next
End Function

Private Function LCFVRslt_VdtDupFld(A As LCFVRslt) As LCFVRslt
Dim O As LCFVRslt: O = A
Dim FldLvs$, J%, IntOpt As IntOpt
Dim OAy() As LCFV, Msg$, Lx2%
For J = 0 To LCFV_UB(A.LCFVAy) - 1
    With LCFVAyIx_DupLxOpt(A.LCFVAy, J)
        If .Som Then
            Lx2 = .Int
            With A.LCFVAy(J)
                Msg = FmtQQ(M_Dup, .Lx, .F, Lx2)
            End With
            Push O.ErLy, Msg
        Else
            LCFV_Push OAy, A.LCFVAy(J)
        End If
    End With
Next
LCFVRslt_VdtDupFld.LCFVAy = OAy
End Function

Private Function LFVAy_VdtErFld(A() As LFV, Fny$()) As LCFVRslt
Dim OAy() As LCFV
Dim OErLy$(), Cno%
Dim J%, M As LCFV, F$, L%, Msg$
For J = 0 To LFV_UB(A)
    With A(J)
        Cno = AyIx(Fny, .F)
        If Cno >= 0 Then
           M.Cno = Cno
           M.F = .F
           M.Lx = .Lx
           M.V = .V
           LCFV_Push OAy, M
        Else
            L = .Lx
            F = .F
            Msg = FmtQQ("Lx(?) Fld(?) is invalid", L, F)
            Push OErLy, Msg
        End If
    End With
Next
With LFVAy_VdtErFld
    .LCFVAy = OAy
    .ErLy = OErLy
End With
End Function

Private Function LVFAy_LFVAy(A() As LVF) As LFV()
Dim O() As LFV, J%, I%, Fny$(), FldLvs$, M As LFV
For J = 0 To LVF_UB(A)
    With A(J)
        M.Lx = .Lx
        M.V = .V
        Fny = LvsSy(.FldLvs)
    End With
    For I = 0 To UB(Fny)
        M.F = Fny(I)
        LFV_Push O, M
    Next
Next
LVFAy_LFVAy = O
End Function

Sub ZZ()
Dim FldAy$(): FldAy = LvsSy("A B C D X W A W")
Dim ValAy$(): ValAy = LvsSy("VA VB VC VD VX VW VA VW")
Dim LxAy%(): LxAy = ApIntAy(10, 20, 30, 40, 50, 60, 70, 80)
Dim A() As LFV, J%
For J = 0 To UB(FldAy)
    LFV_Push3 A, LxAy(J), FldAy(J), ValAy(J)
Next
Dim Fny$(): Fny = LvsSy("B D F A F G H I")
Stop
Dim R As LCFVRslt: R = LFVAy_LCFVRslt(A, Fny)
LCFVRslt_IODmp R
End Sub

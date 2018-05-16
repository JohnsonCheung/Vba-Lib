Attribute VB_Name = "BrkFldValModule"
Option Explicit
Type LFV
    Lx As Integer
    Fld As String
    Val As String
End Type
Type LVF
    Lx As Integer
    Val As String
    FldLvs As String
End Type
Type LFCV
    Lx As Integer
    Fld As String
    Cno As Integer
    Val As String
End Type
Type LFCVRslt
    Itm As String ' Lvl2
    Ay() As LFCV
    ErLy() As String
End Type
Sub LVFAy_Dmp(A() As LVF)
DrsDmp LVFAy_Drs(A)
End Sub
Sub LFCVRslt_Dmp(A As LFCVRslt)
AyDmp LFCVRslt_OupLy(A)
End Sub
Function LVFAy_Drs(A() As LVF) As Drs
With LVFAy_Drs
    .Fny = LvsSy("Lx Val FldLvs")
    .Dry = LVFAy_Dry(A)
End With
End Function
Function LVFAy_Dry(A() As LVF) As Variant()
Dim O(), J%
For J = 0 To LVF_UB(A)
    With A(J)
        Push O, Array(.Lx, .Val, .FldLvs)
    End With
Next
LVFAy_Dry = O
End Function
Function LFCV_Sz%(A() As LFCV)
On Error Resume Next
LFCV_Sz = UBound(A) + 1
End Function
Function LFCV_UB%(A() As LFCV)
LFCV_UB = LFCV_Sz(A) - 1
End Function
Function LFCV_Push(O() As LFCV, A As LFCV)
Dim N%: N = LFCV_Sz(O)
ReDim Preserve O(N)
O(N) = A
End Function

Private Function LFVAy_VdtErFld(A() As LFV, Fny$(), Itm$) As LFCVRslt
Dim OAy() As LFCV
Dim OErLy$(), Cno%
Dim J%, M As LFCV, F$, L%, Msg$
For J = 0 To LFV_UB(A)
    With A(J)
        Cno = AyIx(Fny, .Fld)
        If Cno >= 0 Then
           M.Cno = Cno
           M.Fld = .Fld
           M.Lx = .Lx
           M.Val = .Val
           LFCV_Push OAy, M
        Else
            L = .Lx
            F = .Fld
            Msg = FmtQQ("Lx(?) Fld(?) is invalid", L, F)
            Push OErLy, Msg
        End If
    End With
Next
With LFVAy_VdtErFld
    .Ay = OAy
    .ErLy = OErLy
    .Itm = Itm
End With
End Function
Private Function LFCVRslt_VdtDupFld(A As LFCVRslt) As LFCVRslt
Dim O As LFCVRslt: O = A
Dim FldLvs$, J%, IntOpt As IntOpt
Dim OAy() As LFCV, Msg$, Lx2%
For J = 0 To LFCV_UB(A.Ay) - 1
    With LFCVAyIx_DupLxOpt(A.Ay, J)
        If .Som Then
            Lx2 = .Int
            With A.Ay(J)
                Msg = FmtQQ("Lx(?) Fld(?) is found duplicated in Lx(?).  This item is ignored", .Lx, .Fld, Lx2)
            End With
            Push O.ErLy, Msg
        Else
            LFCV_Push OAy, A.Ay(J)
        End If
    End With
Next
O.Ay = OAy
LFCVRslt_VdtDupFld = O
End Function

Function LFVAy_LFCVRslt(A() As LFV, Fny$(), Itm$) As LFCVRslt
Dim A1 As LFCVRslt: A1 = LFVAy_VdtErFld(A, Fny, Itm)
LFVAy_LFCVRslt = LFCVRslt_VdtDupFld(A1)
End Function

Private Function LFCVAyIx_DupLxOpt(A() As LFCV, Ix%) As IntOpt
Dim I%, F$
F = A(Ix).Fld
For I = Ix + 1 To LFCV_UB(A)
    If F = A(I).Fld Then LFCVAyIx_DupLxOpt = SomInt(A(I).Lx): Exit Function
Next
End Function
Sub Tst()
Dim FldAy$(): FldAy = LvsSy("A B C D X W A W")
Dim ValAy$(): ValAy = LvsSy("VA VB VC VD VX VW VA VW")
Dim LxAy%(): LxAy = ApIntAy(10, 20, 30, 40, 50, 60, 70, 80)
Dim A() As LFV, J%
For J = 0 To UB(FldAy)
    LFV_PushLFV A, LxAy(J), FldAy(J), ValAy(J)
Next
Dim Fny$(): Fny = LvsSy("B D F A F G H I")
Dim R As LFCVRslt: R = LFVAy_LFCVRslt(A, Fny, "Tst")
'LFCVRslt_Dmp R, A, Fny
End Sub
Function ErLy_Dt(ErLy$()) As Dt
ErLy_Dt = AyDt(ErLy, "Er", "Oup ErLy")
End Function
Function LFCVRslt_OupLy(A As LFCVRslt) As String()
LFCVRslt_OupLy = DsLy(LFCVRslt_Ds(A))
End Function

Function LFCVRslt_Ds(A As LFCVRslt) As Ds
Dim O As Ds
O.DsNm = A.Itm & " Rslt"
DsAddDt O, LFCVRslt_Dt(A)
DsAddDt O, ErLy_Dt(A.ErLy)
LFCVRslt_Ds = O
End Function

Function LFCVRslt_VdtValIsLng(A As LFCVRslt) As LFCVRslt
'Dim W%, J%
'Dim O%()
'For J = 0 To UB(IsNumIxAy)
'    W = ValAy(IsNumIxAy(J))
'    If 2 > W Or W > 200 Then Push O, J
'Next
'LFCVRslt_VdtValIsLng = LFCVRslt_New
End Function
Function LFCVAy_FldValLy(A As LFCVRslt, T1$) As String()

End Function
Function LFCVAy_ValFldLy(A As LFCVRslt, T1$) As String()

End Function

Function LFCVRslt_RsltLy(A As LFCVRslt, InpDta() As LFV, InpFny$()) As String()
Dim O$()
Push O, "Itm = " & A.Itm
Push O, "*Inp-Fny ========================================="
PushAy O, InpFny
PushAy O, DtLy(LFVAy_Dt(InpDta))
PushAy O, DsLy(LFCVRslt_Ds(A))
LFCVRslt_RsltLy = O
End Function
Function LFVAy_Dry(A() As LFV) As Variant()
Dim O(), J%
For J = 0 To LFV_UB(A)
    With A(J)
        Push O, Array(.Lx, .Fld, .Val)
    End With
Next
LFVAy_Dry = O
End Function
Function LFVAy_Dt(A() As LFV) As Dt
With LFVAy_Dt
    .Dry = LFVAy_Dry(A)
    .Fny = LvsSy("Lx Fld Val")
    .DtNm = "Lx Fld Val"
End With
End Function

Sub LFV_PushLFV(O() As LFV, Lx%, Fld$, Val$)
Dim M As LFV
With M
    .Fld = Fld
    .Lx = Lx
    .Val = Val
End With
LFV_Push O, M
End Sub
Sub LVF_PushLVF(O() As LVF, Lx%, Val$, FldLvs$)
Dim M As LVF
With M
    .FldLvs = FldLvs
    .Lx = Lx
    .Val = Val
End With
LVF_Push O, M
End Sub
Sub LVF_Push(O() As LVF, A As LVF)
Dim N%: N = LVF_Sz(O)
ReDim Preserve O(N)
O(N) = A
End Sub
Function LVFAy_FldLvsAy(A() As LVF)
Dim O$(), J%
For J = 0 To LVF_UB(A)
    Push O, A(J).FldLvs
Next
LVFAy_FldLvsAy = O
End Function
Function LVF_Sz%(A() As LVF)
On Error Resume Next
LVF_Sz = UBound(A) + 1
End Function
Sub LFV_Push(O() As LFV, A As LFV)
Dim N%: N = LFV_Sz(O)
ReDim Preserve O(N)
O(N) = A
End Sub
Function LFV_Sz%(A() As LFV)
On Error Resume Next
LFV_Sz = UBound(A) + 1
End Function
Function LFV_UB%(A() As LFV)
LFV_UB = LFV_Sz(A) - 1
End Function
Function LVF_UB%(A() As LVF)
LVF_UB = LVF_Sz(A) - 1
End Function
Private Function LVFAy_LFVAy(A() As LVF) As LFV()
Dim O() As LFV, J%, I%, Fny$(), FldLvs$, M As LFV
For J = 0 To LVF_UB(A)
    With A(J)
        M.Lx = .Lx
        M.Val = .Val
        Fny = LvsSy(.FldLvs)
    End With
    For I = 0 To UB(Fny)
        M.Fld = Fny(I)
        LFV_Push O, M
    Next
Next
LVFAy_LFVAy = O
End Function
Function LFCVAy_Fny(A() As LFCV) As String()
Dim O$(), J%
For J = 0 To LFCV_UB(A)
    With A(J)
        PushNoDup O, .Fld
    End With
Next
LFCVAy_Fny = O
End Function

Function LVFAy_Fny(A() As LVF) As String()
Dim O$(), J%
For J = 0 To LVF_UB(A)
    PushAy O, LvsSy(A(J).FldLvs)
Next
LVFAy_Fny = O
End Function
Function LFCVAy_CnoAy(A() As LFCV) As Integer()
Dim O%(), J%
For J = 0 To LFCV_UB(A)
    Push O, A(J).Cno
Next
LFCVAy_CnoAy = O
End Function

Function LVFAy_LFCVRslt(A() As LVF, Fny$(), Itm$) As LFCVRslt
Dim A1() As LFV: A1 = LVFAy_LFVAy(A)
LVFAy_LFCVRslt = LFVAy_LFCVRslt(A1, Fny, Itm)
End Function

Function LFCVRslt_New(A As LFCVRslt, Ay() As LFCV, ErLy$()) As LFCVRslt
Dim O As LFCVRslt: O = A
With O
    .Ay = Ay
    PushAy .ErLy, ErLy
End With
LFCVRslt_New = O
End Function

Function LFCVRslt_VdtIsNum(A As LFCVRslt) As LFCVRslt
Dim OAy() As LFCV
Dim OErLy$(), J%, Msg$
For J = 0 To LFCV_UB(A.Ay)
    With A.Ay(J)
        If IsNum(.Val) Then
            LFCV_Push OAy, A.Ay(J)
        Else
            Msg = FmtQQ("Lx(?) Fld(?) has non-number value(?)", .Lx, .Fld, .Val)
            Push OErLy, Msg
        End If
    End With
Next
LFCVRslt_VdtIsNum = LFCVRslt_New(A, OAy, OErLy)
End Function

Function LFCVRslt_VdtValBet(A As LFCVRslt, FmNum&, ToNum&) As LFCVRslt
Dim V&, J%, Msg$
Dim OErLy$(), OAy() As LFCV
For J = 0 To LFCV_UB(A.Ay)
    With A.Ay(J)
        V = Val(.Val)
        If FmNum > V Or V > ToNum Then
            Msg = FmtQQ("Lx(?) Fld(?) has value(?) outside(?-?)", .Lx, .Fld, .Val, FmNum, ToNum)
            Push OErLy, Msg
        Else
            LFCV_Push OAy, A.Ay(J)
        End If
    End With
Next
LFCVRslt_VdtValBet = LFCVRslt_New(A, OAy, OErLy)
End Function

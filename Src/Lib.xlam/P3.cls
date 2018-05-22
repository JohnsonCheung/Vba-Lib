VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "P3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_Ay() As P3LFV
Private A_L_Lx%()       ' L is for line-record
Private A_L_IsVF() As Boolean
Private Type Er
    Er() As String
    Ay() As P3LCFV
End Type
Private Sub Init(A As LxVFs)
Dim Ay() As LxVF
Ay = A.Ay

End Sub
Sub PushVFAy(A() As VF)
Dim J%
For J = 0 To UB(A)
    PushVF J, A(J)
Next
End Sub
Sub PushVF(Lx%, A As VF)
With A
    If .IsVF Then
        PushLVF Lx, .V, .F
    Else
        PushLFV Lx, .F, .V
    End If
End With
End Sub
Sub PushFVLy(FVLy$())
Dim J%, A$, B$
For J = 0 To UB(FVLy)
    LinAsgTRst FVLy(J), A, B
    PushLFV J, A, B
Next
End Sub

Private Sub ZAddLx(Lx%, IsVF As Boolean)
If Sz(A_L_Lx) = 0 Then
    Push A_L_Lx, Lx
    Push A_L_IsVF, IsVF
Else
    Dim U%: U = UB(A_L_Lx)
    Dim LasLx%: LasLx = A_L_Lx(U)
    If Lx < LasLx Then PrmEr
    If A_L_IsVF(U) <> IsVF Then PrmEr
End If
End Sub
Sub PushLFV(Lx%, F$, V$, Optional IsVF As Boolean)
Dim M As New P3LFV
ZAddLx Lx, IsVF
With M
    .Lx = Lx
    .V = V
    .F = F
End With
PushObj A_Ay, M
End Sub

Sub PushLVF(Lx%, V$, FldLvs$)
Dim M As P3LFV, Fny$()
Fny = LvsSy(FldLvs)
Dim J%
For J = 0 To UB(Fny)
    Set M = New P3LFV
    With M
        .Lx = Lx
        .V = V
        .F = Fny(J)
    End With
    PushObj A_Ay, M
Next
End Sub

Sub PushVFLy(VFLy$())
Dim J%
For J = 0 To UB(VFLy)
    With LinBrkFstTermRst(VFLy(J))
        PushLVF J, .FstTerm, .Rst
    End With
Next
End Sub

Function Vdt(Fny$()) As P3LCFVRslt
Dim O As New P3LCFVRslt
Dim Ay() As P3LCFV
Dim E1 As Er: E1 = ZEr1InvalidFld(Fny)
Dim E2 As Er: E2 = ZEr2DupFld(E1.Ay, Fny)
Dim Er$(): Er = AyAddAp(E1.Er, E2.Er)
Set Vdt = O.Init(E2.Ay, Er)
End Function

Function VdtAsFny() As P3FnyRslt
Dim O As New P3FnyRslt
Dim Fny$(): Fny = Oy.PrpSy("F")
Dim Ok$()
Dim Er$()
Set VdtAsFny = O.Init(Fny, Ok, Er)
End Function
Private Function ZNm$()

End Function
Private Function ZNmEr1_NoLin() As String()

End Function
Private Function ZNmEr2_ExcessLin() As String()

End Function
Function VdtAsNm(T1$, T2$) As P3NmRslt
Dim O As New P3NmRslt
Dim Nm$, Ok$(), Er$()
Dim E1$(): E1 = ZNmEr1_NoLin()
Dim E2$(): E2 = ZNmEr2_ExcessLin()
Nm = ZNm
Er = AyAddAp(E1, E2)
Ok = ApSy(T1 & " " & T2 & " " & Nm)
Set VdtAsNm = O.Init(Nm, Ok, Er)
End Function
Function Show() As String()
End Function
Friend Function ZIOLy(A As LxVFs, Fny$(), FmNum&, ToNum&) As String()
Init A
Dim O$()
     Push O, "P3(======================================================"
PushItmAy O, "Inp1 <<<<<<<<<<<<<<<<<<<<<< :: LxVFs <<<<<<<<<<<<<<<<<<<<", Show
PushItmAy O, "Inp2 <<<<<<<<<<<<<<<<<<<<<< :: Fny   <<<<<<<<<<<<<<<<<<<<", JnSpc(Fny)
PushItmAy O, "Inp3 <<<<<<<<<<<<<<<<<<<<<< :: (FmNum,ToNum) <<<<<<<<<<<<", FmtQQ("(?,?)", FmNum, ToNum)
PushItmAy O, "Oup1=Vdt(Fny) >>>>>>>>>>>>> :: P3LCFVRslt >>>>>>>>>>>>>>>", Vdt(Fny).Show
PushItmAy O, "Oup2=VdtAsNm(Fny) >>>>>>>>> :: P3NmRslt >>>>>>>>>>>>>>>>>", Vdt(Fny).Show
PushItmAy O, "Oup3=VdtAsFny >>>>>>>>>>>>> :: P3FnyRslt >>>>>>>>>>>>>>>>", VdtAsFny.Show
PushItmAy O, "Oup4=VdtAsNumBet(Fny,F,T) > :: P3LCFVRslt >>>>>>>>>>>>>>>", VdtAsNumBet(Fny, FmNum, ToNum).Show
     Push O, "P3)======================================================"
ZIOLy = O
End Function

Function VdtAsNumBet(Fny$(), FmNum&, ToNum&) As P3LCFVRslt
Dim O As New P3LCFVRslt
Dim Ay() As P3LCFV
Dim Er() As String
Stop
Set VdtAsNumBet = O.Init(Ay, Er)
End Function

Private Function ZZInpLy() As String()
Dim O$()
Push O, "Inp-[(Lx,VF)]"
PushAy O, ZZLxVFs.Ly
ZZInpLy = O
End Function

Private Function ZZLxVFAy() As LxVF()
End Function

Private Function ZZLxVFs() As LxVFs
Dim O As New LxVFs
Set ZZLxVFs = O.Init(ZZLxVFAy)
End Function
Friend Function ZZIOLy() As String()
ZZIOLy = ZIOLy(ZZLxVFs, ZZFny, ZZFmNum, ZZToNum)
End Function
Private Function ZZFmNum&()
ZZFmNum = 2
End Function
Private Function ZZToNum&()
ZZToNum = 100
End Function
Private Function ZZFny() As String()
ZZFny = LvsSy("A B C D E F")
End Function
Private Function Oy() As Oy
Dim O As New Oy
Set Oy = O.Init(A_Ay)
End Function

Private Function ZDupIntOpt(J%, Ay() As P3LCFV) As IntOpt
Dim I%, F$
F = Ay(J).Fld
For I = J + 1 To UB(Ay)
    If Ay(I).Fld = F Then
        ZDupIntOpt = SomInt(Ay(I).Lx)
        Exit Function
    End If
Next
End Function

Private Function ZEr1InvalidFld(Fny$()) As Er
Dim Er$(), Ix%
Dim OAy() As P3LCFV
Dim J%, Msg$, M As P3LCFV
For J = 0 To UB(A_Ay)
    With A_Ay(J)
        Ix = AyIx(Fny, A_Ay(J).F)
        If Ix >= 0 Then
            Set M = New P3LCFV
            M.Fld = .F
            M.Lx = .Lx
            M.Cno = Ix
            M.Val = .V
            PushObj OAy, M
        Else
            Msg = FmtQQ(M_Fld_IsInValid, .Lx, .F)
            Push Er, Msg
        End If
    End With
Next
With ZEr1InvalidFld
    .Ay = OAy
    .Er = Er
End With
End Function

Private Function ZEr2DupFld(Ay() As P3LCFV, Fny$()) As Er
Dim Er$(), Msg$, Lx%, F$, J%
Dim OAy() As P3LCFV
For J = 0 To UB(Ay)
    With ZDupIntOpt(J, Ay)
        If .Som Then
            Lx = Ay(J).Lx
            F = Ay(J).Fld
            Msg = FmtQQ(M_Fld_IsDup, Lx, F, .Int)
            Push Er, Msg
        Else
            PushObj OAy, Ay(J)
        End If
    End With
Next
With ZEr2DupFld
    .Er = Er
    .Ay = OAy
End With
End Function

Private Function ZLCFVAy(Ay() As LFV) As LCFV()

End Function

Private Function ZLxAy() As Integer()
ZLxAy = AyUniq(OyPrpIntAy(A_Ay, "Lx"))
End Function

Private Function ZOupDrs() As Drs
ZOupDrs = Vdt(ZZFny).Drs
End Function

Private Function ZZAy() As P3LFV()
Dim O() As P3LFV
Dim M As P3LFV
Set M = New P3LFV: M.F = "AA": M.Lx = 23: M.V = "KLSDFJ": PushObj O, M
Set M = New P3LFV: M.F = "AA": M.Lx = 23: M.V = "KLSDFJ": PushObj O, M
Set M = New P3LFV: M.F = "AA": M.Lx = 23: M.V = "KLSDFJ": PushObj O, M
Set M = New P3LFV: M.F = "AA": M.Lx = 23: M.V = "KLSDFJ": PushObj O, M
ZZAy = O
End Function

Private Function ZZVFLy() As String()
Dim O$()
Push O, "10 A B C D"
Push O, "20 D E F"
ZZVFLy = O
End Function

Friend Sub Tst()
AyDmp ZZIOLy
End Sub
Private Function ZZVFAy() As VF()

End Function

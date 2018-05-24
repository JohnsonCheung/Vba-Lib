VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "P3Itm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_Nm$
Private A_IsVF As Boolean
Private X As New P3SharedFunctions
Private A_Ay() As P3Lin
Private Type Er
    Er() As String
    Ay() As P3LFV
End Type
Property Get IsVF() As Boolean
IsVF = A_IsVF
End Property
Function Init(Nm$, IsVF As Boolean) As P3Itm
A_Nm = Nm
A_IsVF = IsVF
Erase A_Ay
Set Init = Me
End Function

Private Function IsEmp() As Boolean
IsEmp = N = 0
End Function

Property Get N%()
N = Sz(A_Ay)
End Property

Sub Add(Lx%, B$, C$)
Dim O As New P3Lin
With O
    .Lx = Lx
    .B = B
    .C = C
    Set .Par = Me
End With
ZPush A_Ay, O
End Sub
Private Function ZP3LBCAy(A() As P3LFV) As P3LBC()

End Function
Private Function ZP3Itm(A() As P3LFV) As P3Itm
Dim O As P3Itm: Set O = P3Itm(A_Nm, A_IsVF)
Dim A1() As P3LBC, J%
A1 = ZP3LBCAy(A)
For J = 0 To UB(A1)
    With A1(J)
        O.Add .Lx, .B, .C
    End With
Next
Set ZP3Itm = O
End Function
Function Vdt(Fny$()) As P3ItmRslt
If IsEmp Then Set Vdt = New P3ItmRslt: Exit Function
Dim Ay() As P3LFV: Ay = ZLFVAy
Dim E1 As Er: E1 = ZEr1_DupFld(Ay, Fny)
Dim E2 As Er: E2 = ZEr2_InvalidFld(E1.Ay, Fny)
Dim OEr$(): OEr = AyAddAp(E1.Er, E2.Er)
Set Vdt = P3ItmRslt(ZP3Itm(E2.Ay), OEr)
End Function

Function VdtAsFny() As P3FnyRslt
Dim O As New P3FnyRslt
Dim Fny$(): Fny = ZLFVOy.PrpSy("F")
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
Private Function ZOneLinLCFVAy(Lx%) As LCFV()
ZOneLinLCFVAy = Oy.WhPrpEqVal("Lx", Lx)
End Function
Sub AddBCAy(BCAy$())
Dim Lx%, B$, C$
For Lx = 0 To UB(BCAy)
    With Brk(BCAy(Lx), " ")
        Add Lx, .S1, .S2
    End With
Next
End Sub
Friend Function ZIOLy(Nm$, IsVF As Boolean, BCAy$(), Fny$(), FmNum&, ToNum&) As String()
Init Nm, IsVF
AddBCAy BCAy
Dim O$()
    ZPush O, "P3(======================================================"
PushItmAy O, "Inp1 <<<<<<<<<<<<<<<<<<<<<< :: LxVFs <<<<<<<<<<<<<<<<<<<<", Show
   PushAp O, "Inp2 <<<<<<<<<<<<<<<<<<<<<< :: Fny   <<<<<<<<<<<<<<<<<<<<", JnSpc(Fny)
   PushAp O, "Inp3 <<<<<<<<<<<<<<<<<<<<<< :: (FmNum,ToNum) <<<<<<<<<<<<", FmtQQ("(?,?)", FmNum, ToNum)
PushItmAy O, "Oup1=Vdt(Fny) >>>>>>>>>>>>> :: P3ItmRslt >>>>>>>>>>>>>>>>", Vdt(Fny).Show
PushItmAy O, "Oup2=VdtAsNm(Fny) >>>>>>>>> :: P3NmRslt >>>>>>>>>>>>>>>>>", Vdt(Fny).Show
PushItmAy O, "Oup3=VdtAsFny >>>>>>>>>>>>> :: P3FnyRslt >>>>>>>>>>>>>>>>", VdtAsFny.Show
PushItmAy O, "Oup4=VdtAsNumBet(Fny,F,T) > :: P3ItmRslt >>>>>>>>>>>>>>>>", VdtAsNumBet(Fny, FmNum, ToNum).Show
    ZPush O, "P3)======================================================"
ZIOLy = O
End Function
Private Function Cv(I) As P3Lin
Set Cv = I
End Function
Private Function ZLFVOy() As Oy
Set ZLFVOy = Vb.Oy(ZLFVAy)
End Function
Private Function ZP3LFV(Lx%, F$, V$) As P3LFV
Dim O As New P3LFV
With O
    .F = F
    .Lx = Lx
    .V = V
End With
Set ZP3LFV = O
End Function
Private Property Get ZLFVAy() As P3LFV()
If IsEmp Then Exit Property
Dim O() As P3LFV, Lx%, I, B$, C$, F, M As P3Lin
For Each I In A_Ay
    Set M = I
    B = M.B
    C = M.C
    Lx = M.Lx
    If A_IsVF Then
        For Each F In LvsSy(C)
            PushObj O, ZP3LFV(Lx, CStr(F), B)
        Next
    Else
        PushObj O, ZP3LFV(Lx, B, C)
    End If
Next
ZLFVAy = O
End Property

Function VdtAsNumBet(Fny$(), FmNum&, ToNum&) As P3ItmRslt
Dim O As New P3ItmRslt
If IsEmp Then Set VdtAsNumBet = O: Exit Function
Dim Ay() As P3Lin
Dim Er() As String
Dim I, V&, Msg$, M As P3LFV
For Each I In ZLFVAy
    Set M = I
    With M
        If IsNum(.V) Then
            V = .V
            If FmNum <= V And V <= ToNum Then
                PushObj Ay, I
            Else
                Msg = FmtQQ(M_Val_ShouldBet, .Lx, .F, .V, FmNum, ToNum)
                ZPush Er, Msg
            End If
        Else
            Msg = FmtQQ(M_Val_IsNonNum, .Lx, .F, .V)
            ZPush Er, Msg
        End If
    End With
Next
Stop
'Set VdtAsNumBet = O.Init(ZP3Itm(A), Er)
End Function

Private Function ZP3Lin(OLx%, Lin$, IsVF As Boolean) As P3Lin
Set ZP3Lin = X.LxLin_P3Lin(OLx, Lin, IsVF)
OLx = OLx + 1
End Function
Private Sub ZZResBCLy()
'10 A B C
'20 D E
'30 A A E
End Sub
Private Function ZZBCLy() As String()
ZZBCLy = MdResLy(IdeMd.Md("P3Itm"), "BCLy")
End Function
Private Function ZZNm$()
ZZNm = "Wdt"
End Function
Private Function ZZIsVF() As Boolean
ZZIsVF = True
End Function

Friend Function ZZIOLy() As String()
ZZIOLy = ZIOLy(ZZNm, ZZIsVF, ZZBCLy, ZZFny, ZZFmNum, ZZToNum)
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

Private Function ZDupIntOpt(J%, Ay() As P3LFV) As IntOpt
Dim I%, F$
'F = Ay(J).Fld
For I = J + 1 To UB(Ay)
    If Ay(I).F = F Then
        ZDupIntOpt = SomInt(Ay(I).Lx)
        Exit Function
    End If
Next
End Function

Private Function ZEr2_InvalidFld(A() As P3LFV, Fny$()) As Er
If AyIsEmp(A) Then Exit Function
Dim Er$(), Ix%
Dim OAy() As P3LFV
Dim J%, Msg$, M As P3LFV
For J = 0 To UB(A)
    With A(J)
        Ix = AyIx(Fny, A(J).F)
        If Ix >= 0 Then
            Set M = New P3LFV
            M.F = .F
            M.Lx = .Lx
            M.V = .V
            PushObj OAy, M
        Else
            Msg = FmtQQ(M_Fld_IsInValid, .Lx, .F)
            ZPush Er, Msg
        End If
    End With
Next
With ZEr2_InvalidFld
    .Ay = OAy
    .Er = Er
End With
End Function

Private Function ZEr1_DupFld(A() As P3LFV, Fny$()) As Er
Dim Er$(), Msg$, Lx%, F$, J%
Dim OAy() As P3LFV
For J = 0 To UB(A)
    With ZDupIntOpt(J, A)
        If .Som Then
            Lx = A(J).Lx
            F = A(J).F
            Msg = FmtQQ(M_Fld_IsDup, Lx, F, .Int)
            ZPush Er, Msg
        Else
            PushObj OAy, A(J)
        End If
    End With
Next
With ZEr1_DupFld
    .Er = Er
    .Ay = OAy
End With
End Function


Private Function ZLxAy() As Integer()
ZLxAy = AyUniq(OyPrpIntAy(A_Ay, "Lx"))
End Function

Property Get Nm$()
Nm = A_Nm
End Property
Friend Sub Tst()
AyDmp ZZIOLy
End Sub
Property Get Fny() As String()
Dim O$(), J%
Dim Ay() As P3Lin: 'Ay = A_Lins.Ay
For J = 0 To UB(Ay)
    PushNoDupAy O, Ay(J).Fny
Next
Fny = O
End Property
Private Sub ZPush(O, A)
VbAy.Push O, A
End Sub
Function Show() As String()
Dim O$(), J%
ZPush O, "P3Itm("
ZPush O, FmtQQ("Nm,IsVF=[?],[?]", A_Nm, A_IsVF)
For J = 0 To U
    PushAy O, A_Ay(J).Show
Next
ZPush O, "P3Itm)"
Show = O
End Function
Private Function U%()
U = UB(A_Ay)
End Function

Property Get Ly() As String()
Dim O$(), J%
For J = 0 To U
    ZPush O, A_Ay(J).Lin
Next
Ly = O
End Property
Sub PushLin(Lx%, B$, C$)
Stop
End Sub
Friend Sub Brw()
AyBrw Show
End Sub

Private Function Oy() As Oy
Set Oy = Vb.Oy(A_Ay)
End Function

Private Function ZZEr() As String()
ZZEr = ApSy("Lx(3) Is error")
End Function

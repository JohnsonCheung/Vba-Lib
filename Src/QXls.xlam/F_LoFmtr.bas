Attribute VB_Name = "F_LoFmtr"
'Option Explicit
'Public Const C1_Lo$ = "Lo"
'Public Const C1_Fmt$ = "Fmt"
'Public Const C1_Wdt$ = "Wdt"
'Public Const C1_Lvl$ = "Lvl"
'Public Const C1_Cor$ = "Cor"
'Public Const C1_Bdr$ = "Bdr"
'Public Const C1_Tit$ = "Tit"
'Public Const C1_Lbl$ = "Lbl"
'Public Const C1_Fml$ = "Fml"
'Public Const C1_Bet$ = "Bet"
'Public Const C1_Tot$ = "Tot"
'Public Const C2_Tot_Sum$ = "Sum"
'Public Const C2_Tot_Avg$ = "Avg"
'Public Const C2_Tot_Cnt$ = "Cnt"
'Public Const C2_Lo_Nam$ = "Nam"
'Public Const C2_Lo_Fld$ = "Fld"
'Public Const C2_Lo_Hid$ = "Hid"
'Public Const C2_Bdr_L$ = "Left"
'Public Const C2_Bdr_R$ = "Right"
'Public Const C2_Bdr_C$ = "Col"
'Public Const M_Fld_IsAvg_FndInSum$ = "Lx(?) has Fld(?), which is TAvg-Fld, but also found in TSum-Lx(?)"
'Public Const M_Fld_IsCnt_FndInSum$ = "Lx(?) has Fld(?), which is TCnt-Fld, but also found in TSum-Lx(?)"
'Public Const M_Fld_IsCnt_FndInAvg$ = "Lx(?) has Fld(?), which is TCnt-Fld, but also found in TAvg-Lx(?)"
'Public Const M_Val_IsNonNum$ = "Lx(?) has Val(?) should be a number"
'Public Const M_Val_IsNonLng$ = "Lx(?) has Val(?) should be a 'Long' number"
'Public Const M_Val_ShouldBet$ = "Lx(?) has Val(?) should be between [?] and [?]"
'Public Const M_Fld_IsInValid$ = "Lx(?) Fld(?) is invalid.  Not found in Fny"
'Public Const M_Fld_IsDup$ = "Lx(?) Fld(?) is found dup in Lx(?)"
'Public Const M_Nm_LinHasNoVal$ = "Lx(?) is Nm-Lin, it has no value"
'Public Const M_Nm_NoNmLin$ = "Nm-Lin is missing"
'Public Const M_Nm_ExcessLin$ = "LX(?) is excess due to Nm-Lin is found above"
'Public Const M_Should_Lng$ = "Lx(?) Fld(?) should have val(?) be a long number"
'Public Const M_Should_Num$ = "Lx(?) Fld(?) should have val(?) be a number"
'Public Const M_Should_Bet$ = "Lx(?) Fld(?) should have val(?) be between (?) and (?)"
'Public Const M_Bet_Should2Term = "Lx(?) Fld(?) is Bet-Line.  It should have 2 terms"
'Public Const M_Bet_InvalidTerm = "Lx(?) Fld(?) is Bet-Line.  It has invalid term(?)"
'Public Const M_Fld_Dup$ = "Lx(?) Fld(?) is found duplicated in Lx(?).  This item is ignored"
'
'Private Type BdrInp
'    BdrL() As LABC
'    BdrR() As LABC
'    BdrC() As LABC
'End Type
'Private Type TotInp
'    Sum() As LABC
'    Avg() As LABC
'    Cnt() As LABC
'End Type
'Private Type VF 'Val FldLvs
'     Nam() As LABC
'     Hid() As LABC
'     Fny() As LABC
'     Bdr As BdrInp
'     Cor() As LABC
'     Fmt() As LABC
'     Lvl() As LABC
'     Tot As TotInp
'     Wdt() As LABC
'End Type
'Private Type FV
'     Bet() As LABC
'     Fml() As LABC
'     Lbl() As LABC
'     Tit() As LABC
'End Type
'Private Type TotRslt
'    SumCny() As String
'    AvgCny() As String
'    CntCny() As String
'    ErLy() As String
'    ABCLy() As String
'End Type
'Private Type BdrRslt
'    CnoAy() As Integer
'    ABCLy() As String
'    ErLy() As String
'End Type
'Private Type Inp
'    VF As VF
'    FV As FV
'End Type
'Private Type CnoStr
'    Cno() As Integer
'    Str() As String
'End Type
'Private Type CnoLng
'    Cno() As Integer
'    Lng() As Long
'End Type
'Private Type Bet5
'    BetC() As Integer
'    BetA() As String
'    BetB() As String
'End Type
'Private A_Ly$()
'
'Private Sub ZZ_LoFmtrLy_Rslt()
'A_Ly = ZZLy
'
''Dim B$
''    B = "Bdr"
''Dim Ay() 'Ay = \B = Ay = Array(ToStr, Tag("Wdt", Validate.FmtWs.Wdt.ToStr))
''    Dim S$
''        Dim O
''        Dim Pth$
''        Pth = FmtQQ("?.ToStr", B)
''        Set O = CallByName(Validate.FmtWs, B, VbGet)
''        S = O.ToStr
''    Ay = Array(ToStr, Tag(B, S))
''AyBrw Array(ToStr, Validate.ToStr)
'Stop '
'End Sub
'
'Private Function LABCAyRslt_VdtBet(A As LABCAyRslt, Fny$()) As LABCAyRslt
'Stop '
''If A.LABCs.IsVF Then PmEr
''Dim OEr  As Er: OEr.Push A.Er
''Dim O As LABCs
''Dim HasEr As Boolean
''    Set O = A.LABCs.DupEmpLABCs
''    Dim A1 As LABCs: Set A1 = A.LABCs
''    Dim J%
''    HasEr = True
''    While HasEr
''        J = J + 1
''        If J > 1000 Then Stop
''        ZVdtBet_1 A1, Fny, _
''            O, OEr, HasEr
''        Set A1 = O
''    Wend
''Set LABCAyRslt_VdtBet = LABCAyRslt(O, OEr)
'End Function
'
'Private Function ZVdtFml(A As LABCAyRslt, Fny$()) As LABCAyRslt
'Stop '
'''If A.LABCs.IsVF Then PmEr
''Dim Er$(): PushAy Er, A.ErLy
''Dim O() As LABC
''    Stop
'''    Set O = A.LABCs.DupEmpLABCs
''    Dim A1 As LABCs: Set A1 = A.LABCs
''    Dim HasEr As Boolean
''    Dim J%
''    HasEr = True
''    While HasEr
''        J = J + 1
''        If J > 1000 Then Stop
''        ZVdtFml_1 A1, _
''            O, OEr, HasEr
''        Set A1 = O
''    Wend
''Set ZVdtFml = LABCAyRslt(O, OEr)
'End Function
'
'Private Sub ZVdtFml_1(A() As LABC, _
'    O() As LABC, OErLy$(), OHasEr As Boolean)
'Stop '
'''Each LABC in A,
'''   if there is error in the Fml,
'''      then Push Er  to OEr
'''      else Push Itm to O & set OHasEr = true
''Set O = A.DupEmpLABCs
''OHasEr = False
''    Dim F$()
''    Dim M As SomStr
''    Dim J%, Ay() As LABC
''    F = A.UniqFny
''    Ay = A.Ay
''    For J = 0 To UB(Ay)
''        With Ay(J)
''            M = Fml(.B).ErMsgOpt(F)
''            If M.Som Then
''                OHasEr = True           '<==
''                OEr.PushMsg M.Str       '<==
''            Else
''                O.AddLBC .Lx, .B, .C    '<==
''            End If
''        End With
''    Next
'End Sub
'
'Private Sub ZVdtBet_1(A() As LABC, Fny$(), _
'    O() As LABC, OErLy$(), OHasEr As Boolean)
'Stop '
''Set O = A.DupEmpLABCs
''OHasEr = False
''    Dim ErLy$()
''    Dim F$(): F = A.UniqFny
''    Dim J%, Ay() As LABC
''    Ay = A.Ay
''    For J = 0 To UB(Ay)
''        With Ay(J)
''            ErLy = ZVdtBet_2(.C, .Lx, .FldLvs, Fny)
''            If AyIsEmp(ErLy) Then
''                O.AddLBC .Lx, .B, .C '<==
''            Else
''                OHasEr = True       '<==
''                OEr.PushErLy0 ErLy   '<==
''            End If
''        End With
''    Next
'End Sub
'Private Function ZVdtBet_2(C$, Lx%, F$, Fny$()) As String()
''C$ is the col-c of Bet-line.  It should have 2 item and in Fny
''Return Er of M_Bet_* if any
'Dim A$()
'    A = SslSy(C)
'If Sz(A) <> 2 Then
'    ZVdtBet_2 = ApSy(FmtQQ(M_Bet_Should2Term, Lx, F))
'    Exit Function
'End If
'If Not AyHas(Fny, A(0)) Then
'    ZVdtBet_2 = ApSy(FmtQQ(M_Bet_InvalidTerm, Lx, F, A(0)))
'    Exit Function
'End If
'If Not AyHas(Fny, A(1)) Then
'    ZVdtBet_2 = ApSy(FmtQQ(M_Bet_InvalidTerm, Lx, F, A(1)))
'    Exit Function
'End If
'
'End Function
'
'Private Function ZVdtTot(A As TotInp, Fny$()) As TotRslt
'Stop '
''Dim Avg1 As LABCAyRslt: Set Avg1 = A.Avg.ValidateAsFldVal(Fny)
''Dim Sum1 As LABCAyRslt: Set Sum1 = A.Sum.ValidateAsFldVal(Fny)
''Dim Cnt1 As LABCAyRslt: Set Cnt1 = A.Cnt.ValidateAsFldVal(Fny)
''
''Dim A1() As LxFld, S1() As LxFld, C1() As LxFld
''    A1 = Avg1.LABCs.LxFldAy
''    S1 = Sum1.LABCs.LxFldAy
''    C1 = Cnt1.LABCs.LxFldAy
''Dim A2$(), S2$(), C2$()
''Dim A3%(), S3%(), C3%()
''    With Oy(A1)
''        A2 = .PrpSy("Fld")
''        A3 = .PrpIntAy("Lx")
''    End With
''    With Oy(S1)
''        S2 = .PrpSy("Fld")
''        S3 = .PrpIntAy("Lx")
''    End With
''    With Oy(C1)
''        C2 = .PrpSy("Fld")
''        C3 = .PrpIntAy("Lx")
''    End With
''Dim Avg$(), Cnt$(), Sum$()
''    Sum = S2
''    Avg = AyMinus(A2, S2)
''    Cnt = AyMinusAp(C2, S2, A2)
''Dim Er As Er
''    Dim E1 As Er, E2 As Er
''    Set E1 = ZVdtTot_1(C2, C3, S2, S3, A2, A3)
''    Set E2 = ZVdtTot_2(A2, A3, S2, S3)
''    Er.PushAp Avg1.Er, Sum1.Er, Cnt1.Er, E1, E2
''Dim Ly$()
''    Dim L$
''    If Sz(Sum) > 0 Then
''        L = JnSpc(Sum)
''        L = FmtQQ("? ? ?", C1_Tot, C2_Tot_Sum, L)
''        Push Ly, L
''    End If
''    If Sz(Avg) > 0 Then
''        L = JnSpc(Avg)
''        L = FmtQQ("? ? ?", C1_Tot, C2_Tot_Avg, L)
''        Push Ly, L
''    End If
''    If Sz(Cnt) > 0 Then
''        L = JnSpc(Cnt)
''        L = FmtQQ("? ? ?", C1_Tot, C2_Tot_Cnt, L)
''        Push Ly, L
''    End If
''Dim O As TotRslt
''    O.AvgCny = Avg
''    O.CntCny = Cnt
''    O.SumCny = Sum
''    O.ABCLy = Ly
''    Set O.Er = Er
''ZVdtTot = O
'End Function
'
'Private Function ZVdtTot_1(Cnt$(), CntLxAy%(), Sum$(), SumLxAy%(), Avg$(), AvgLxAy%()) As String()
'Stop '
''Dim O As Er
''Dim J%, C$, Ix%, Msg$
''For J = 0 To UB(Cnt)
''    C = Cnt(J)
''    Ix = AyIx(Sum, C)
''    If Ix >= 0 Then
''        Msg = FmtQQ(M_Fld_IsCnt_FndInSum, CntLxAy(J), Cnt(J), SumLxAy(Ix))
''        O.PushMsg Msg
''    Else
''        Ix = AyIx(Avg, C)
''        If Ix >= 0 Then
''            Msg = FmtQQ(M_Fld_IsCnt_FndInAvg, CntLxAy(J), Cnt(J), AvgLxAy(Ix))
''            O.PushMsg Msg
''        End If
''    End If
''Next
''Set ZVdtTot_1 = O
'End Function
'Private Function ZVdtTot_2(Avg$(), AvgLxAy%(), Sum$(), SumLxAy%()) As String()
'Stop '
''Dim O As Er
''Dim J%, A$, Ix%, Msg$
''For J = 0 To UB(Avg)
''    A = Avg(J)
''    Ix = AyIx(Sum, A)
''    If Ix >= 0 Then
''        Msg = FmtQQ(M_Fld_IsAvg_FndInSum, AvgLxAy(J), Avg(J), SumLxAy(Ix))
''        O.PushMsg Msg
''    End If
''Next
''Set ZVdtTot_2 = O
'End Function
'Private Function ZVdtBdr(A As BdrInp, Fny$()) As BdrRslt
'Stop '
''Dim LL As LABCAyRslt: Set LL = A.BdrL.ValidateAsFldVal(Fny)
''Dim RR As LABCAyRslt: Set RR = A.BdrR.ValidateAsFldVal(Fny)
''Dim CC As LABCAyRslt: Set CC = A.BdrC.ValidateAsFldVal(Fny)
''Dim O As BdrRslt
''    Dim B%() 'B is Left-Bdr-CnoAy to be return
''    Dim C$() 'C is Ly of [Lo Bdr? ..]
''    Dim L1$(): L1 = LL.LABCs.UniqFny 'L1 is Left-Fny of input after LABCs.Validate
''    Dim R1$(): R1 = RR.LABCs.UniqFny 'R1 is Right-Fny ..
''    Dim C1$(): C1 = CC.LABCs.UniqFny 'C1 is Col-Fny ..
''    B = ZVdtBdr_1(L1, R1, C1, Fny)
''    C = ZVdtBdr_2(B, Fny)           'C is Ly of [Lo BdrL ..
''                                    '                  C
''                                    '                  R ..  <- only one line or no line of R.
''    With O
''        .CnoAy = B
''        .ABCLy = C
''        .Er.PushAp LL.Er, RR.Er, CC.Er
''    End With
''ZVdtBdr = O
'End Function
'
'Private Function ZVdtBdr_3(LeftCnoAy%()) As Integer()
'Dim O%()
'    Dim J%, U%
'    U = UB(LeftCnoAy)
'    For J = 0 To U Step 2
'        If J = U Then Exit For
'        If LeftCnoAy(J) + 1 = LeftCnoAy(J + 1) Then Push O, LeftCnoAy(J)
'    Next
'ZVdtBdr_3 = O
'End Function
'
'Private Function ZVdtBdr_2(LeftCnoAy%(), Fny$()) As String()
'Dim L%()    ' L is Left-CnoAy coming from LeftCnoAy
'Dim C%()    ' C is Col-CnoAy  coming from LeftCnoAy
'    Dim Rst%()
'    Dim CC%()
'    Dim J%
'    Rst = AySrt(LeftCnoAy)
'Again:
'    J = J + 1
'    If J > 1000 Then Stop
'    CC = ZVdtBdr_3(Rst)
'    Rst = ZVdtBdr_5(Rst, CC) ' Remove those columns (Left&Right) of CC from Rst (They are all Left)
'    If Sz(CC) > 0 Then
'        PushAy C, CC
'        GoTo Again
'    End If
'    L = Rst
'Dim R$      ' R is the FldNm should set the Right-Bdr.
'    R = ZVdtBdr_4(L, C, Fny, _
'        L, C)
'Dim L1$   'Left-Bdr-FldLvs
'    Dim Sy$()
'    '\L =
'    For J = 0 To UB(L)
'        Push Sy, Fny(L(J))
'    Next
'    L1 = JnSpc(Sy)
'Dim C1$   'Column-Bdr-FldLvs
'    '\C =
'    Erase Sy
'    For J = 0 To UB(C)
'        Push Sy, Fny(C(J))
'    Next
'    C1 = JnSpc(Sy)
'Dim O$() '\L1 C1 R =
'    Push O, FmtQQ("? ? ?", C1_Bdr, C2_Bdr_L, L1)
'    Push O, FmtQQ("? ? ?", C1_Bdr, C2_Bdr_C, C1)
'    If R <> "" Then
'        Push O, FmtQQ("? ? ?", C1_Bdr, C2_Bdr_R, R)
'    End If
'ZVdtBdr_2 = O
'End Function
'Private Function ZVdtBdr_4$(L%(), C%(), Fny$(), _
'    OL%(), OC%())
''Return R$ as Right-Bdr-FldNm or Blank
''L% is Left-Bdr-CnoAy in ascending order
''C% is Col-Bdr-CnoAy in ascending order
''R will be Las-Ele-of-Fny
''   only if L has value and Las-ele-of-L is Sz(Fny) (case 1)
''        or last-ele-of-C <> UB(Fny)                (case 2)
''if (case 2), OL = AyRmvLasEle(L)
''if (case 1), OC = AyRmvLasEle(C)
'Dim N%: N = Sz(Fny)
'If Sz(L) > 0 Then
'    If AyLasEle(L) = N Then     'should compare with N
'        ZVdtBdr_4 = AyLasEle(Fny)
'        OL = AyRmvLasEle(L)
'        OC = C
'        Exit Function
'    End If
'End If
'If Sz(C) > 0 Then
'    If AyLasEle(L) = N - 1 Then     'Should compare with N-1
'        ZVdtBdr_4 = AyLasEle(Fny)
'        OL = L
'        OC = AyRmvLasEle(C)
'        Exit Function
'    End If
'End If
'OL = L
'OC = C
'End Function
'
'Private Function ZVdtBdr_5(L%(), C%()) As Integer()
''Return L - C
'Dim CC%(), J%
'    For J = 0 To UB(C)
'        Push CC, C(J)
'        Push CC, C(J) + 1
'    Next
'ZVdtBdr_5 = AyMinus(L, CC)
'End Function
'
'Private Function ZVdtBdr_1(L$(), R$(), C$(), Fny$()) As Integer()
''Inp-L: Left-Bdr-Fny
''    R: Right-Bdr-Fny
''    C: Column-Bdr-Fny
''Ret: Left-Bdr-CnoAy
'Dim F$(), Cno&()
'Dim J%
'
'Dim RR%()
'    F = AyAdd(C, R)              ' F is those column with either Col-Bdr or Right-Bdr
'    Cno = AyIxAy(Fny, F)         ' Cno is Cno-of-Fny1
'    For J = 0 To UB(Cno)
'        If Cno(J) < 0 Then Stop  ' the validate of InvalidFld has problem
'        Push RR, Cno(J) + 1      ' Add 1 to each Cno to become R, so that R becomes Left-Bdr-Cno
'    Next
'Dim LL%()
'    F = AyAdd(C, L)             ' F is those column with either Col-Bdr or Left-Bdr
'    Cno = AyIxAy(Fny, F)        ' Cno is Cno-of-Fny
'    For J = 0 To UB(Cno)
'        If Cno(J) < 0 Then Stop ' the validate of invalidFld has problem
'        Push LL, Cno(J) + 1     ' Put the Cno to to L
'    Next
'Dim O%()
'    PushNoDupAy O, LL
'    PushNoDupAy O, RR
'ZVdtBdr_1 = O
'End Function
'
'Private Function ZInp() As Inp
'Stop '
''Dim I As Inp
''With I
''    .FV.Bet.InitByT1 C1_Bet
''    .FV.Fml.InitByT1 C1_Fml
''    .FV.Lbl.InitByT1 C1_Lbl
''    .FV.Tit.InitByT1 C1_Tit
''    .VF.Nam.InitByT1 C1_Lo, IsVF:=True
''    .VF.Hid.InitByT1 C1_Lo, IsVF:=True
''    .VF.Fny.InitByT1 C1_Lo, IsVF:=True
''    .VF.Cor.InitByT1 C1_Cor, IsVF:=True
''    .VF.Fmt.InitByT1 C1_Fmt, IsVF:=True
''    .VF.Lvl.InitByT1 C1_Lvl, IsVF:=True
''    .VF.Wdt.InitByT1 C1_Wdt, IsVF:=True
''    .VF.Tot.Cnt.InitByT1 C1_Tot, IsVF:=True
''    .VF.Tot.Avg.InitByT1 C1_Tot, IsVF:=True
''    .VF.Tot.Sum.InitByT1 C1_Tot, IsVF:=True
''    .VF.Bdr.BdrL.InitByT1 C1_Bdr, IsVF:=True
''    .VF.Bdr.BdrC.InitByT1 C1_Bdr, IsVF:=True
''    .VF.Bdr.BdrR.InitByT1 C1_Bdr, IsVF:=True
''End With
''Dim J%, A$, B$, C$
''Dim Er$()
''For J = 0 To UB(A_Ly)
''    Linx(A_Ly(J)).AsgTTRst A, B, C
''    Select Case A
''    Case C1_Lo
''        Select Case B
''            Case C2_Lo_Hid: I.VF.Hid.AddLBC J, B, C
''            Case C2_Lo_Fld: I.VF.Fny.AddLBC J, B, C
''            Case C2_Lo_Nam: I.VF.Nam.AddLBC J, B, C
''        End Select
''    Case C1_Bdr
''        Select Case B
''            Case C2_Bdr_L: I.VF.Bdr.BdrL.AddLBC J, B, C
''            Case C2_Bdr_R: I.VF.Bdr.BdrR.AddLBC J, B, C
''            Case C2_Bdr_C: I.VF.Bdr.BdrC.AddLBC J, B, C
''            Case Else: Push Er, FmtQQ("Lx(?) T2(?) should be [Left Right Col]", J, B)
''        End Select
''    Case C1_Tot
''        Select Case B
''            Case C2_Tot_Sum: I.VF.Tot.Sum.AddLBC J, B, C
''            Case C2_Tot_Avg: I.VF.Tot.Avg.AddLBC J, B, C
''            Case C2_Tot_Cnt: I.VF.Tot.Cnt.AddLBC J, B, C
''            Case Else:  Push Er, FmtQQ("Lx(?) T2(?) should be [Tot Avg Cnt]", J, B)
''        End Select
''    Case C1_Fmt: I.VF.Fmt.AddLBC J, B, C
''    Case C1_Fml: I.FV.Fml.AddLBC J, B, C
''    Case C1_Bet: I.FV.Bet.AddLBC J, B, C
''    Case C1_Tit: I.FV.Tit.AddLBC J, B, C
''    Case C1_Lbl: I.FV.Lbl.AddLBC J, B, C
''    Case C1_Wdt: I.VF.Wdt.AddLBC J, B, C
''    Case C1_Lvl: I.VF.Lvl.AddLBC J, B, C
''    Case C1_Cor: I.VF.Cor.AddLBC J, B, C
''    Case Else
''        Push Er, FmtQQ("Lx(?) T1(?) should be [Lo Wdt Lbl ...]", J, A)
''    End Select
''Next
''ZInp = I
'End Function
'
'Property Get LoFmtrLy_Rslt(A$()) As LoFmtrRslt
'Stop '
''Dim I As Inp:                 I = ZInp
''Dim NamR As NmRslt:    Set NamR = I.VF.Nam.ValidateAsNm
''Dim FnyR As FnyRslt:   Set FnyR = I.VF.Fny.ValidateAsFny(C1_Lo, C2_Lo_Fld)
''Dim Fny$():                 Fny = FnyR.Fny
''Dim BdrR As BdrRslt:       BdrR = ZVdtBdr(I.VF.Bdr, Fny)
''Dim TotR As TotRslt:       TotR = ZVdtTot(I.VF.Tot, Fny)
''Dim FmlR As LABCAyRslt: Set FmlR = ZVdtFml(I.FV.Fml.ValidateAsFldVal(Fny), Fny)
''Dim BetR As LABCAyRslt: Set BetR = ZVdtBet(I.FV.Bet.ValidateAsFldVal(Fny), Fny)
''Dim CorR As LABCAyRslt: Set CorR = I.VF.Cor.ValidateAsFldLngVal(Fny)
''Dim FmtR As LABCAyRslt: Set FmtR = I.VF.Fmt.ValidateAsFldVal(Fny)
''Dim HidR As LABCAyRslt: Set HidR = I.VF.Hid.ValidateAsFldVal(Fny)
''Dim LblR As LABCAyRslt: Set LblR = I.FV.Lbl.ValidateAsFldVal(Fny)
''Dim LvlR As LABCAyRslt: Set LvlR = I.VF.Lvl.ValidateAsBetNum(Fny, 2, 6)
''Dim TitR As LABCAyRslt: Set TitR = I.FV.Tit.ValidateAsFldVal(Fny)
''Dim WdtR As LABCAyRslt: Set WdtR = I.VF.Wdt.ValidateAsBetNum(Fny, 2, 100)
''Dim OFmtWs As FmtWs
''Dim OEr As Er
''Dim Ok$()
''    OEr.PushAp BdrR.Er, BetR.Er, CorR.Er, FmlR.Er, HidR.Er, LblR.Er, LvlR.Er, TitR.Er, TotR.Er, WdtR.Er
''    Ok = Ly0Ap_Ly( _
''        NamR.Lin, _
''        FnyR.Lin, _
''        LvlR.LABCs.Ly, _
''        BdrR.ABCLy, _
''        BetR.LABCs.Ly, _
''        CorR.LABCs.Ly, _
''        FmlR.LABCs.Ly, _
''        FmtR.LABCs.Ly, _
''        HidR.LABCs.Ly, _
''        LblR.LABCs.Ly, _
''        TitR.LABCs.Ly, _
''        TotR.ABCLy, _
''        WdtR.LABCs.Ly)
''    With OFmtWs
''        .SetBdr BdrR.CnoAy
''        .SetTot TotR.SumCny, TotR.AvgCny, TotR.CntCny
''        .SetHid HidR.LABCs.CnoVals(Fny).CnoAy
''        .SetWdt WdtR.LABCs.CnoVals(Fny)
''        .SetBet BetR.LABCs.CnoVals(Fny)
''        .SetCor CorR.LABCs.CnoVals(Fny)
''        .SetFml FmlR.LABCs.CnoVals(Fny)
''        .SetFmt FmtR.LABCs.CnoVals(Fny)
''        .SetLvl LvlR.LABCs.CnoVals(Fny)
''        .SetTit TitR.LABCs.CnoVals(Fny)
''        .SetLbl LblR.LABCs.CnoVals(Fny)
''    End With
''Dim O As LoFmtrRslt
''Set Validate = O.Init(Ok, OEr, OFmtWs)
'End Property
'
'Private Sub ZZ_ZZLy()
'AyDmp ZZLy
'End Sub
'
'Private Function ZZLy() As String()
'ZZLy = ResNm_Ly("Ly")
'End Function
'
'Private Sub ZZRes_Ly()
''Lo Nam ABC
''Lo Fld A B C D E F G
''Lo Hid B C X
''Bet A C D
''Bet B D E
''Wdt 10 A B X
''Wdt 20 D C C
''Wdt 3000 E F G C
''Fmt #,## A B C
''Fmt #,##.## D E
''Lvl 2 A C
''Bdr Left A
''Bdr Right G
''Bdr Col F
''Tot Sum A B
''Tot Cnt C
''Tot Avg D
''Tit A abc | sdf
''Tit B abc | sdkf | sdfdf
''Cor 12345 A B
''Fml F A + B
''Fml C A * 2
''Lbl A lksd flks dfj
''Lbl B lsdkf lksdf klsdj f
'End Sub
'
'Sub LoFmtrTp_Brw()
'Dim O$()
'Push O, "Lo  Nm     *Nm"
'Push O, "Lo  Fld    *Fld.."
'Push O, "Lo  Hid    *Fld.."
'Push O, "Bdr Left   *Fld.."
'Push O, "Bdr Right  *Fld.."
'Push O, "Bdr Col    *Fld.."
'Push O, "Tot Tot    *Fld.."
'Push O, "Tot Avg    *Fld.."
'Push O, "Tot Cnt    *Fld.."
'
'Push O, "Fmt *Fmt   *Fld.."
'Push O, "Wdt *Wdt   *Fld.."
'Push O, "Lvl *Lvl   *Fld.."
'Push O, "Cor *Cor   *Fld.."
'
'Push O, "Tit *Fld   *Tit"
'Push O, "Lbl *Fld   *Lbl"
'Push O, "Fml *Fld   *Formula"
'Push O, "Bet *Fld   *Fld1 *Fld2"
'AyBrw O
'End Sub
'
'
'Sub LoFmtr_FmtLo(A As LoFmtr, Lo As ListObject)
'
'End Sub
'Private Sub ZZ_DoFmt()
'Dim A As New LoFmtr
'Dim B As LoFmtrRslt
''Set B = A.InitBySampleLy.Validate
''B.FmtWs.DoFmt SampleLo
''A.InitBySampleLy.Validate.FmtWs.DoFmt SampleLo
'End Sub
'
'Sub DoFmt(A As ListObject)
'DoFmtBdr A
'DoFmtFmt A
'DoFmtCor A
'DoFmtHid A
'DoFmtFml A
'DoFmtLvl A
'DoFmtBet A
'DoFmtTit A
'DoFmtTot A
'DoFmtLbl A
'End Sub
'
'Sub DoFmtBet(A As ListObject)
'Stop '
''Dim Ay() As CnoVal, C%, V, J%, Fml$
''Ay = A_Bet.Ay
''For J = 0 To UB(Ay)
''    C = Ay(J).Cno
''    With Brk(Ay(J).V, " ")
''        Fml = FmtQQ("=Sum([?]:[?])", .S1, .S2)
''    End With
''    LoC(A, C).Formula = Fml
''Next
'End Sub
'Sub DoFmtBdr(A As ListObject)
'Stop '
''Dim J%, C%(): C = A_BdrCnoAy
''For J = 0 To UB(C)
''    RgBdrLeft LoC(A, C(J))
''Next
'End Sub
'Sub DoFmtFmt(A As ListObject)
'Stop '
''Dim Ay() As CnoVal, C%, V, J%
''Ay = A_Fmt.Ay
''For J = 0 To UB(Ay)
''    C = Ay(J).Cno
''    V = Ay(J).V
''    LoC(A, C).NumberFormat = V
''Next
'End Sub
'Sub DoFmtCor(A As ListObject)
'Stop '
''Dim Ay() As CnoVal, C%, V, J%
''Ay = A_Cor.Ay
''For J = 0 To UB(Ay)
''    C = Ay(J).Cno
''    V = Ay(J).V
''    LoC(A, C).Interior.Color = V
''Next
'End Sub
'Sub DoFmtFml(A As ListObject)
'Stop '
''Dim Ay() As CnoVal, C%, V, J%
''Ay = A_Fml.Ay
''For J = 0 To UB(A_Fml)
''    C = Ay(J).Cno
''    V = Ay(J).V
''    LoC(A, C).Formula = V
''Next
'End Sub
'Sub DoFmtHid(A As ListObject)
'Stop '
''Dim J%, C%(): C = A_HidCnoAy
''For J = 0 To UB(C)
''    LoC(A, C(J)).EntireColumn.Hidden = True
''Next
'End Sub
'Sub DoFmtLvl(A As ListObject)
'Stop '
''LoWs(A).Outline.SummaryColumn = xlSummaryOnLeft
''Dim Ay() As CnoVal, C%, V, J%
''Ay = A_Lvl.Ay
''For J = 0 To UB(Ay)
''    C = Ay(J).Cno
''    V = Ay(J).V
''    LoC(A, C).EntireColumn.OutlineLevel = V
''Next
'End Sub
'
'Sub DoFmtTot(A As ListObject)
'Stop '
''Y A, A_SumCny, xlTotalsCalculationSum
''Y A, A_AvgCny, xlTotalsCalculationAverage
''Y A, A_CntCny, xlTotalsCalculationCount
'End Sub
'
'Sub DoFmtWdt(A As ListObject)
'Stop '
''Dim Ay() As CnoVal, C%, V, J%
''Ay = A_Wdt.Ay
''For J = 0 To UB(Ay)
''    C = Ay(J).Cno
''    V = Ay(J).V
''    LoC(A, C).ColumnWidth = V
''Next
'End Sub
'
'Private Function ZTitNRow%()
'Stop '
''Dim O%
''    Dim A$(), J%, M%
''    A$() = A_Tit.StrValAy
''    For J = 0 To UB(A)
''        M = Sz(Split(A(J), "|"))
''        If M > 0 Then O = M
''    Next
''ZTitNRow = O
'End Function
'
'Sub DoFmtTit(A As ListObject)
'Dim Fny$():         Fny = LoFny(A)
'Dim A1 As Range: Set A1 = A.DataBodyRange
'Dim HasTit As Boolean ' HasTit means if there is enough space above DtaRg to put the title
'    Dim N%: N = ZTitNRow
'    Dim TitR%: TitR = A1.Row - 2 - N ' Tit-Row
'    HasTit = TitR >= 0
'If Not HasTit Then Exit Sub
'
'Dim At As Range         ' At is TitAt
'    Set At = WsRC(RgWs(A), TitR, A1.Column) ' Use Ws to find this 'At', because TitR is relative to Ws
'    Dim Sq(): Sq = ZTitSq(Fny)
'TitRg_Fmt SqRg(Sq, At)   '<-- put the title and fmt the title Range
'End Sub
'
'Sub DoFmtLbl(A As ListObject)
'Stop '
''Dim Lbl$, Ay() As CnoVal, J%, C%
''Dim RFld As Range
''Dim RLbl As Range
''Ay = A_Lbl.Ay
''For J = 0 To UB(Ay)
''    C = Ay(J).Cno
''    Lbl = Ay(J).V
''    Set RFld = RgRC(LoC(A, C), 0, 1)
''    Set RLbl = RgRC(RFld, 0, 1)
''    RLbl.Value = RFld.Value    '<-- Swapping
''    RFld.Value = Lbl    '<-- Swapping
''Next
'End Sub
'Private Function ZTitDry(Fny$()) As Variant()
'Stop '
'''From A_Tit & Fny, return TitDry
'''If some column has no title, use FldNm as Tit
''Dim O(), J%, Ix%, Ay() As CnoVal
''For J = 0 To UB(Fny)
''    Ix = A_Tit.CnoIx(J)
''    If Ix = -1 Then
''        Push O, Array(Fny(J))
''    Else
''        Push O, SplitVBar(Ay(J).V, Trim:=True)  ' V contains Tit of current fld
''    End If
''Next
''ZTitDry = O
'End Function
'
'Private Function ZTitSq(Fny$()) As Variant()
'Stop
''Dim TitSq(): TitSq = DrySq(ZTitDry(Fny))
''ZTitSq = SqTranspose(TitSq)
'End Function
'Private Sub Y(Lo As ListObject, C$(), A As XlTotalsCalculation)
'Dim J%
'For J = 0 To UB(C)
'    Lo.ListColumns(C(J)).TotalsCalculation = A
'Next
'End Sub
'
'
'
'

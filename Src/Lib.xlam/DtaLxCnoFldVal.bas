Attribute VB_Name = "DtaLxCnoFldVal"
Option Explicit
Type LABC
    Lx As Integer
    A As String
    B As String
    C As String
End Type
Type LCFV
    Lx As Integer
    Cno As Integer
    F As String
    V As Variant
End Type
Type LCFVRslt
    Ok() As String
    LCFVAy() As LCFV
    Er() As String
End Type
Type NmRslt
    Ok() As String
    Er() As String
    Nm As String
End Type
Type FnyRslt
    Ok() As String
    Er() As String
    Fny() As String
End Type
Const M_Should_Lng$ = "Lx(?) Fld(?) should have val(?) be a long number"
Const M_Should_Num$ = "Lx(?) Fld(?) should have val(?) be a number"
Const M_Should_Bet$ = "Lx(?) Fld(?) should have val(?) be between (?) and (?)"
Const M_Dup$ = _
                      "Lx(?) Fld(?) is found duplicated in Lx(?).  This item is ignored"

Sub AAA()
ZZ
End Sub

Function LABCAy_FnyRslt(A() As LABC) As FnyRslt

End Function

Function LABCAy_LCFVRslt(A() As LABC, Fny$(), Optional IsVF As Boolean) As LCFVRslt
If LABC_Sz(A) = 0 Then Exit Function
Dim OEr$()
Dim A1() As LABC: A1 = ZErOf_InvalidFld(A, Fny, IsVF, OEr) 'All fld in A1 (at B or C by IsVF) are now all valid
                                                           'Any any error is reported in OEr
Dim A2() As LABC: A2 = ZErOf_DupFld(A1, IsVF, OEr)         'In A1, any dup fld is removed, A2 is clean, OEr is cummulated
Dim A3$():        A3 = ZLABCAy_Ly(A2)
Dim A4() As LCFV: A4 = ZLABCAy_LCFVAy(A2, Fny, IsVF)
With LABCAy_LCFVRslt
    .Ok = A3
    .LCFVAy = A4
    .Er = OEr
End With
End Function
Private Function ZErOf_BetNum(A() As LABC, FmNum&, ToNum&, OEr$()) As LABC()
'Assume LABC is always are IsVF=true

End Function
Function LABCAy_LCFVRslt_OfBetNum(A() As LABC, Fny$(), FmNum&, ToNum&) As LCFVRslt
Dim OEr$(), IsVF As Boolean 'Always a IsVF=True line
IsVF = True
Dim A1() As LABC: A1 = ZErOf_InvalidFld(A, Fny, IsVF, OEr)
Dim A2() As LABC: A2 = ZErOf_DupFld(A1, IsVF, OEr)
Dim A3() As LABC: A3 = ZErOf_BetNum(A2, FmNum, ToNum, OEr)
Dim A4() As LCFV: A4 = ZLABCAy_LCFVAy(A2, Fny, IsVF)
Dim A5$():        A5 = ZLABCAy_Ly(A3)
With LABCAy_LCFVRslt_OfBetNum
    .Ok = A5
    .LCFVAy = A4
    .Er = OEr
End With
End Function

Function LABCAy_NmRslt(A() As LABC) As NmRslt
Dim Nm$, Er$(), Ok$()
With LABCAy_NmRslt
    .Er = Er
    .Ok = Ok
    .Nm = Nm
End With
End Function

Function LABC_Push(O() As LABC, A As LABC)
Dim N%: N = LABC_Sz(O)
ReDim Preserve O(N)
O(N) = A
End Function

Function LABC_UB%(A() As LABC)
LABC_UB = LABC_Sz(A) - 1
End Function

Function LCFV_UB%(A() As LCFV)
LCFV_UB = LCFV_Sz(A) - 1
End Function

Private Function LABC_Sz%(A() As LABC)
On Error Resume Next
LABC_Sz = UBound(A) + 1
End Function

Private Sub LCFV_Push(O() As LCFV, A As LCFV)
Dim N%: N = LCFV_Sz(O)
ReDim Preserve O(N)
O(N) = A
End Sub

Private Sub LCFV_PushAy(O() As LCFV, A() As LCFV)
Dim J%
For J = 0 To LCFV_UB(A)
    LCFV_Push O, A(J)
Next
End Sub

Private Function LCFV_Sz%(A() As LCFV)
On Error Resume Next
LCFV_Sz = UBound(A) + 1
End Function

Private Function ZABCLy_LABCAy(Ly$()) As LABC()
Dim O() As LABC, J%
For J = 0 To UB(Ly)
    LABC_Push O, ZLinLABC(Ly(J), J)
Next
ZABCLy_LABCAy = O
End Function

Private Function ZErOf_DupFld(A() As LABC, IsVF As Boolean, OEr$()) As LABC()
Dim J%, O() As LABC, FldLvs$, M As LABC
If IsVF Then
    For J = 0 To LABC_UB(A)
        FldLvs = ""
        ZErOf_DupFld_1 A, J, _
            FldLvs, OEr
        If FldLvs <> "" Then
            M = A(J)
            M.C = FldLvs
            LABC_Push O, M
        End If
    Next
    ZErOf_DupFld = O
    Exit Function
End If
Stop
End Function
Private Function ZErOf_DupFld_1IsDup(Fny$(), I%) As Boolean
'Check if Fny(I)-element has duplication found in Fny(I+1..)
Dim F$: F = Fny(I)
Dim II%
For II = I + 1 To UB(Fny)
    If Fny(II) = F Then ZErOf_DupFld_1IsDup = True: Exit Function
Next
End Function
Private Function ZErOf_DupFld_2IsDup(A() As LABC, J%, F$, _
    ODupAtLx%) As Boolean
'Check if F has duplicated-element found in A(J+1...)
Dim JJ%, Fny$(), FldLvs$
ODupAtLx = -1
For JJ = J + 1 To LABC_UB(A)
    FldLvs = A(JJ).C
    Fny = LvsSy(FldLvs)
    If AyHas(Fny, F) Then
        ODupAtLx = JJ
        ZErOf_DupFld_2IsDup = True
        Exit Function
    End If
Next
End Function

Private Sub ZErOf_DupFld_1(A() As LABC, J%, _
    OFldLvs$, OEr$())
'A(J).C is FldLvs
'Check If it has any duplicated element is same line or following lines
'if yes, remove it by assigning the no-dup-Fny1 to OFldLvs
'        and report the duplication in OEr
OFldLvs = ""
Dim Fny$()
    Dim FldLvs$
    FldLvs = A(J).C
    Fny = LvsSy(FldLvs) '<--- All Fld to be checked
Dim Fny1$() '<-- Those Fld has no duplicated
    Dim I%
    Dim F$
    For I = 0 To UB(Fny)
        F = Fny(I)
        Dim IsDup As Boolean
        Dim DupAtLx% ' Duplicated at which line
            IsDup = ZErOf_DupFld_1IsDup(Fny, I)
            If IsDup Then
                DupAtLx = J 'At same line
            Else
                IsDup = ZErOf_DupFld_2IsDup(A, J, F, _
                    DupAtLx) 'DupAtLx at following line of J+1...
            End If
        
        If IsDup Then
            Dim Lx%, Msg$
            Lx = A(J).Lx
            Msg = FmtQQ(M_Fld_IsDup, Lx, F, DupAtLx)
            Push OEr, Msg   '<== Report duplication in OEr
        Else
            Push Fny1, F    '<== Push to Fny1 for no-dup
        End If
    Next
OFldLvs = JnSpc(Fny1)
End Sub

Private Function ZErOf_InvalidFld(A() As LABC, Fny$(), IsVF As Boolean, OEr$()) As LABC()
Dim J%, O() As LABC, FldLvs$, M As LABC, Lx%, C$, F$
If IsVF Then
    For J = 0 To LABC_UB(A)
        With A(J)
            Lx = .Lx
            C = .C
        End With

        ZErOf_InvalidFld_1 Lx, C, Fny, _
            FldLvs, OEr
                
        If FldLvs <> "" Then
            M = A(J)
            M.C = FldLvs
            LABC_Push O, M  '<============
        End If
    Next
    ZErOf_InvalidFld = O
    Exit Function
End If
Dim Msg$
For J = 0 To LABC_UB(A)
    With A(J)
        Lx = .Lx
        F = .B
    End With
    If AyHas(Fny, F) Then
        LABC_Push O, A(J)       '<======
    Else
        Msg = FmtQQ(M_Fld_IsInValid, Lx, F)
        Push OEr, Msg           '<======
    End If
Next
ZErOf_InvalidFld = O
End Function

Private Sub ZErOf_InvalidFld_1(Lx%, FldLvs$, Fny$(), _
    OFldLvs$, OEr$())
Dim Fny1$(), Fny2$(), J%, F$, Msg$
Fny1 = LvsSy(FldLvs)
For J = 0 To UB(Fny1)
    F = Fny1(J)
    If Not AyHas(Fny, F) Then
        Msg = FmtQQ(M_Fld_IsInValid, Lx, F)
        Push OEr, Msg   '<===================
    Else
        Push Fny2, F
    End If
Next
OFldLvs = JnSpc(Fny2)
End Sub

Private Function ZIOLy1(IsVF As Boolean, ABCLy$(), Fny$()) As String()
Dim I1$(), I2$(), I3$()
Dim O1$(), O2$(), O3$()
Dim LABCAy() As LABC
Dim O$()
Dim R As LCFVRslt

    
LABCAy = ZABCLy_LABCAy(ABCLy)
R = LABCAy_LCFVRslt(LABCAy, Fny, IsVF) '<========================
    I1 = ZLABCAy_Ly(LABCAy)
    I2 = ApSy(JnSpc(Fny))
    I3 = ApSy(IsVF)
    O3 = R.Er
    O1 = R.Ok
    O2 = ZLCFVAy_Ly(R.LCFVAy)

         Push O, "LABCAy_LCFVRslt=============================="
    PushItmAy O, "Inp1::LABCAy <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<", I1
    PushItmAy O, "Inp2::Fny <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<", I2
    PushItmAy O, "Inp3::IsVF <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<", I3
    PushItmAy O, "Oup1::OK >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>", O1
    PushItmAy O, "Oup2::LCFVAy >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>", O2
    PushItmAy O, "Oup3::Er >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>", O3
         Push O, "LABCAy_LCFVRslt)============================="
         Push O, ""
ZIOLy1 = O
End Function

Private Function ZIOLy2(IsVF As Boolean, ABCLy$(), Fny$(), FmNum&, ToNum&) As String()
Dim LABCAy() As LABC
Dim R As LCFVRslt
Dim I1$(), I2$(), I3$()
Dim O1$(), O2$(), O3$()
Dim O$()
    LABCAy = ZABCLy_LABCAy(ABCLy)
    R = LABCAy_LCFVRslt_OfBetNum(LABCAy, Fny, FmNum, ToNum) '<========================
    
    I1 = ZLABCAy_Ly(LABCAy)
    I2 = ApSy(JnSpc(Fny))
    I3 = ApSy(IsVF)
    O1 = R.Ok
    O2 = ZLCFVAy_Ly(R.LCFVAy)
    O3 = R.Er
     
         Push O, "LABCAy_LCFVRsltOfBetNum======================"
    PushItmAy O, "Inp1::LABCAy <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<", I1
    PushItmAy O, "Inp2::Fny <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<", I2
    PushItmAy O, "Inp3::FmNum ToNum <<<<<<<<<<<<<<<<<<<<<<<<<<<", I3
    PushItmAy O, "Oup1::Ok >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>", O1
    PushItmAy O, "Oup2::LCFVAy_OfBetNum >>>>>>>>>>>>>>>>>>>>>>>", O2
    PushItmAy O, "Oup3::Er >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>", O3
         Push O, "LABCAy_LCFVRsltOfBetNum======================"
         Push O, ""
ZIOLy2 = O
End Function

Private Function ZIOLy3(ABCLy$())
Dim LABCAy() As LABC
Dim R1 As NmRslt
Dim O$()
Dim I1$(), O1$(), O2$(), O3$()
    LABCAy = ZABCLy_LABCAy(ABCLy)
    R1 = LABCAy_NmRslt(LABCAy) '<========================
    
    I1 = ZLABCAy_Ly(LABCAy)
    O1 = ApSy(R1.Nm)
    O2 = R1.Ok
    O3 = R1.Er
         
         Push O, "LABCAy_NmRslt ==============================="
    PushItmAy O, "Inp1::LABCAy <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<", I1
    PushItmAy O, "Oup1::Nm >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>", O1
    PushItmAy O, "Oup2::LABCAy Nm >>>>>>>>>>>>>>>>>>>>>>>>>>>>>", O2
    PushItmAy O, "Oup3::Er >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>", O3
         Push O, "LABCAy_NmRslt ==============================="
         Push O, ""
ZIOLy3 = O
End Function

Private Function ZIOLy4(ABCLy$()) As String()
Dim LABCAy() As LABC
Dim R As FnyRslt
Dim O$()
Dim I1$()
    LABCAy = ZABCLy_LABCAy(ABCLy)
    R = LABCAy_FnyRslt(LABCAy) '<============
    
    I1 = ZLABCAy_Ly(LABCAy)
         Push O, "LABCAy_FnyRslt =============================="
    PushItmAy O, "Inp1::LABCAy <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<", I1
    PushItmAy O, "Oup1::Fny >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>", R.Fny
    PushItmAy O, "Oup2::Ok::String[]>>>>>>>>>>>>>>>>>>>>>>>>>>>", R.Ok
    PushItmAy O, "Oup3::Er >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>", R.Er
         Push O, "LABCAy_FnyRslt =============================="
         Push O, ""
ZIOLy4 = O
End Function

Private Function ZLABCAy_LCFVAy(A() As LABC, Fny$(), IsVF As Boolean) As LCFV()
Dim J%, O() As LCFV
For J = 0 To LABC_UB(A)
    With A(J)
        If IsVF Then
                    
        Else
            
        End If
    End With
Next
End Function

Private Function ZLABCAy_Ly(A() As LABC) As String()
Dim J%, O$()
For J = 0 To LABC_UB(A)
    Push O, ZLABC_Lin(A(J))
Next
ZLABCAy_Ly = O
End Function

Private Function ZLABC_Lin$(A As LABC)
With A
    ZLABC_Lin = FmtQQ("? ? ? ?", .Lx, .A, .B, .C)
End With
End Function

Private Function ZLCFVAy_LABCAy(A() As LCFV, T1$, IsVF As Boolean) As LABC()
Dim O() As LABC, M As LABC
If IsVF Then
    ZLCFVAy_LABCAy = ZLCFVAy_LABCAy_1(A, T1)
    Exit Function
End If
End Function

Private Function ZLCFVAy_LABCAy_1(A() As LCFV, T1$) As LABC()
Dim J%, Lx%, M As LABC, LxAy%(), V$, FldLvs$, O() As LABC
LxAy = ZLxAy(A)
For J = 0 To UB(LxAy)
Next
With M
    .A = T1
    .B = V
    .C = FldLvs
    .Lx = Lx
End With
LABC_Push O, M
End Function

Private Function ZLCFVAy_Ly(A() As LCFV) As String()
Dim J%, O$()
For J = 0 To LCFV_UB(A)
    Push O, ZLCFV_Lin(A(J))
Next
ZLCFVAy_Ly = O
End Function

Private Function ZLCFV_Lin$(A As LCFV)
With A
    ZLCFV_Lin = FmtQQ("? ? ? ?", .Lx, .Cno, .F, .V)
End With
End Function

Private Function ZLinLABC(Lin$, Lx%) As LABC
Dim A$, B$, C$
LinAsgTTRst Lin, A, B, C
With ZLinLABC
    .Lx = Lx
    .A = A
    .B = B
    .C = C
End With
End Function

Private Function ZLxAy(A() As LCFV) As Integer()

End Function

Private Sub ZZ()
AyDmp ZZIOLy1
End Sub

Private Function ZZIOLy1() As String()
Dim IsVF As Boolean, ABCLy$(), Fny$()
IsVF = True
Push ABCLy, "Wdt 10 A B C X"
Push ABCLy, "Wdt 20 A B D Y A"
Fny = LvsSy("A B C D E X")
ZZIOLy1 = ZIOLy1(IsVF, ABCLy, Fny)
End Function

Private Function ZZIOLy2() As String()
Dim IsVF As Boolean, ABCLy$(), Fny$()
IsVF = True
Push ABCLy, "Wdt 10 A B C"
Push ABCLy, "Wdt 20 A B D"
Fny = LvsSy("A B C D E")
Dim FmNum&, ToNum&
ZZIOLy2 = ZIOLy2(IsVF, ABCLy, Fny, FmNum, ToNum)
End Function

Private Function ZZIOLy3() As String()
Dim R1 As NmRslt

End Function

Private Function ZZIOLy4() As String()
Dim R2 As FnyRslt

End Function

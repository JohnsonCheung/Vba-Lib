VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SwLnx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private B1_Lx As Integer
Private B1_Nm As String
Private B1_OpStr As String
Private B1_TermAy() As String
Private B1_Pm As Dictionary
Private Type SomBool
    Bool As Boolean
    Som As Boolean
End Type
Private Type SomV
    V As Variant
    Som As Boolean
End Type
Property Get AndOrLinEr$()
If Not BoolOpStr_IsAndOr(B_OpStr) Then Exit Property
Dim O$
    Dim NTerm%
    NTerm = Sz(B_TermAy)

If NTerm < 2 Then
    O = "" ' 'Msg = "" B_Lx, "For OR|AND, must have 2 or more operands"): GoTo Ext
    Stop
    GoTo X
    Exit Property
End If
Dim T1$: T1 = B_TermAy(0)
Dim T2$: T2 = B_TermAy(1)

Select Case FstChr(T1)
    Case "%"
        If Not B_Pm.Exists(T1) Then
            O = "" '(B_Lx, "For OR|AND, first term must be found in Pm")
            Stop
            GoTo X
        End If
    Case Else
        O = "" ' NewTpEr(Lx, "For OR|AND, first operand must begin with %")
        Stop
        GoTo X
    End Select

Select Case FstChr(T2)
    Case "%"
        If Not B_Pm.Exists(T2) Then
            O = "" ' NewTpEr(Lx, "For EQ|NE, second operand not found in Pm")
            Stop
            GoTo X
        End If
    Case "?"
        O = "" 'O = NewTpEr(Lx, "For EQ|NE, second operand cannot begin with ?")
        Stop
        GoTo X
    Case "*"
        If UCase(T1) <> "*BLANK" Then
            O = "" ' NewTpEr(Lx, "For AND|OR, second operand can be *BLANK, but nothing else begin with *")
            Stop
            GoTo X
        End If
    End Select
X:
    AndOrLinEr = O
End Property
Private Function T1T2_SomBool(T1$, T2$, EqNeOpStr$, Sw As Dictionary) As SomBool
Dim S1$, S2$
    With Term_VarOpt(T1, Sw)
        If Not .Som Then Exit Function
        S1 = .V
    End With
    With Term_VarOpt(T2, Sw)
        If Not .Som Then Exit Function
        S2 = .V
    End With
Dim O As SomBool
Select Case EqNeOpStr
Case "EQ": O = SomBool(S1 = S2)
Case "NE": O = SomBool(S1 <> S2)
Case Else: Stop
End Select
T1T2_SomBool = O
End Function
Private Function Term_VarOpt(A, Sw As Dictionary) As SomV
'switch-term begins with % or ? or it is *Blank.  % is for parameter & ? is for switch
'  If %, it will evaluated to str by lookup from Pm
'        if not exist in {Pm}, stop, it means the validation fail to remove this term
'  If ?, it will evaluated to bool by lookup from Sw
'        if not exist in {Sw}, return None
'  Otherwise, just return SomVar(A)
Dim O As VarOpt
    Select Case FstChr(A)
    Case "?"
        If Not Sw.Exists(A) Then Exit Function
        O = SomVar(Sw(A))
    Case "%"
        If Not B_Pm.Exists(A) Then
            Stop ' it means the validation fail to remove this term
        End If
        O = SomVar(B_Pm(A))
    Case "*"
        If A <> "*Blank" Then Stop ' it means the validation fail to remove this term
        O = SomVar("")
    Case Else
        O = SomVar(A)
    End Select
Term_VarOpt = O
End Function

Private Function TermAy_SomBool(A$(), Op, Sw As Dictionary) As SomBool ' Op as  As e_BoolAyOp
Dim B As New Bools
    Dim I
    For Each I In A
        With Term_VarOpt(I, Sw)
            If Not .Som Then Exit Function
            B.Push CBool(.V)
        End With
    Next
TermAy_SomBool = SomBool(B.Val(Op))
End Function

Private Function SomBool(Sw As Dictionary) As SomBool
If Sw.Exists(B_Nm) Then Exit Function
Dim Ay$(): Ay = B_TermAy
Dim Z As SomBool
Select Case B_OpStr
Case "OR":  Z = TermAy_SomBool(Ay, "OR", Sw)
Case "AND": Z = TermAy_SomBool(Ay, "AND", Sw)
Case "NE":  Z = T1T2_SomBool(Ay(0), Ay(1), "NE", Sw)
Case "EQ":  Z = T1T2_SomBool(Ay(0), Ay(1), "EQ", Sw)
Case Else: Stop
End Select
If Not Z.Som Then Exit Function
SomBool = SomBool(Z.Bool)
End Function

Private Property Get EqNeLinEr$()
If Not BoolOpStr_IsEqNe(B_OpStr) Then Exit Property
Dim O$
    Dim NTerm%
    Dim Msg$
    NTerm = Sz(B_TermAy)
If NTerm <> 2 Then
    Stop
    O = "" 'O = NewTpEr(Lx, "When 2nd-Term (Operator) is [AND OR], only 2 terms are allowed")
    GoTo X
End If
X:
    EqNeLinEr = O
End Property
Property Get Nm$()
Nm = B_Nm
End Property
Property Get Lx%()
Lx = B_Lx
End Property
Property Get ErLy() As String()
ErLy = Ly0Ap_Ly(NmEr, OpEr, PfxEr, EqNeLinEr, AndOrLinEr)
End Property

Private Property Get NmEr$()
If B_Nm = "" Then
    Stop
    NmEr = "" '  A.Lx, "The line has no name")
End If
End Property

Private Property Get OpEr$()
If BoolOpStr_IsVdt(B_OpStr) Then Exit Property
OpEr = "" ' NewTpEr(A.Lx, "Invalid operator.  Valid operation [NE EQ AND OR]")
Stop
End Property

Private Property Get PfxEr$()
If FstChr(B_Nm) <> "?" Then
    PfxEr = "" ' NewTpEr(A.Lx, "First char must be [?]")
    Stop
End If
End Property

Private Function TermEr() As String()
Dim O0$(), O1$(), O2$(), I
Dim FldNmSet As Dictionary
Dim TermAy$()
Dim O$()
Dim Lx%
For Each I In TermAy
    If HasPfx(I, "?") Then
        If Not FldNmSet.Exists(I) Then Push O, I
    ElseIf HasPfx(I, "%?") Then
        If Not B_Pm.Exists(I) Then Push O, I
    Else
        Push O2, I
    End If
Next
Dim B$, C$
If Not AyIsEmp(O0) Then C = FmtQQ("[?] must be found in Switch", JnSpc(O0))
If Not AyIsEmp(O1) Then C = FmtQQ("[?] must be found in Pm", JnSpc(O1))
If Not AyIsEmp(O2) Then B = FmtQQ("[?] must begin with [ ? | %? ]", JnSpc(O1))
Dim Sy$()
    PushNonEmp Sy, C
    PushNonEmp Sy, B
If Sz(Sy) > 0 Then
    Stop
    'O = NewTpEr(Lx, JnCrLf(Sy))
End If
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SwLnxs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type Evl      ' During each cycle of evaluation, return this Type
    Ay() As SwLnx     ' the SwLnxAy to be evaluated
    Sw As Dictionary  ' the cummulated Sw after evaluated
    HasEvl As Boolean ' will be true, if there is some SwLnx evaluated
End Type
Private Type SwLnxAyErLy
    SwLnxAy() As SwLnx
    ErLy() As String
End Type
Private Type SwErLy
    Sw As Sw
    ErLy() As String
End Type
Private B_Ay() As SwLnx
Private B_Pm As Dictionary

Property Get Init(A() As Lnx, Pm As Dictionary) As SwLnxs
Dim M As SwLnx, J%
For J = 0 To UB(A)
    Set M = New SwLnx
    PushObj B_Ay, M.Init(A(J), Pm)
Next
Set B_Pm = Pm
Set Init = Me
End Property

Property Get Rslt() As SwRslt
Dim R1 As SwLnxAyErLy
Dim R2 As SwLnxAyErLy
Dim R3 As SwErLy
    R1 = Rslt_1
    R2 = DupNmEr(R1.SwLnxAy)
    R3 = Rslt_2(R2.SwLnxAy)
Dim E1$(), E2$(), E3$()
    E1 = R1.ErLy
    E2 = R2.ErLy
    E3 = R3.ErLy
Dim E$()
    E = Ly0Ap_Ly(E1, E2, E3)
Dim O As New SwRslt
With O
    .ErLy = E
    Set .Sw = R3.Sw
End With
Set Rslt = O
End Property

Friend Sub Tst()
StrBrw Tag_NyStr_ObjAp("SwLnxsIO", Me, Rslt)
End Sub

Private Property Get DupNmEr(A() As SwLnx) As SwLnxAyErLy
Dim Ny$(): Ny = DupNmEr_0(A)
Dim DupNy$(): DupNy = AyDupAy(Ny)
Dim DupNmLx%(): DupNmLx = DupNmEr_1(A, DupNy)
Dim E$(): E = DupNmEr_2(DupNmLx, A)
Dim Ay() As SwLnx: Ay = DupNmEr_3(DupNmLx)
With DupNmEr
    .SwLnxAy = Ay
    .ErLy = E
End With
End Property

Private Function DupNmEr_0(A() As SwLnx) As String()
Dim O$(), J%
For J = 0 To U
    Push O, A(J).Nm
Next
DupNmEr_0 = O
End Function

Private Function DupNmEr_1(A() As SwLnx, DupNy$()) As Integer()
End Function

Private Function DupNmEr_2(DupNmLx%(), A() As SwLnx) As String()
Dim O$()
    Dim Msg$, J%
    For J = 0 To UB(DupNmLx)
        Dim Nm$, Lx%
        Lx = DupNmLx(J)
        Nm = DupNmEr_4(Lx, A)
        Msg = FmtQQ(M_, Lx, Nm)
        Stop
        Push O, Msg
    Next
DupNmEr_2 = O
End Function

Private Function DupNmEr_3(Lx%()) As SwLnx()
Dim O() As SwLnx
Dim J%
For J = 0 To U
    If Not AyHas(Lx, B_Ay(J).Lx) Then
        PushObj O, B_Ay(J)
    End If
Next
DupNmEr_3 = O
End Function

Private Function DupNmEr_4$(Lx%, A() As SwLnx)
'return A(Lx=Lx).Nm
Dim J%
For J = 0 To UB(A)
    If A(J).Lx = Lx Then DupNmEr_4 = A(J).Nm: Exit Function
Next
Stop ' LgcEr
End Function

Private Property Get N&()
N = UB(B_Ay)
End Property

Private Function Rslt_1() As SwLnxAyErLy
'Return a split-of-SwLnxAy-and-ErLy as SwLnxAyErLy
'by if B_Ay(..).ErLy has Er
'       then put into ErLy    (E$())
'       else put into SwLnxAy (O() As SwLnx)
Dim E$(), O() As SwLnx
Dim J%, Er$()
For J = 0 To U
    Er = B_Ay(J).ErLy
    If AyIsEmp(Er) Then
        PushObj O, B_Ay(J)
    Else
        PushAy E, Er
    End If
Next
With Rslt_1
    .ErLy = E
    .SwLnxAy = O
End With
End Function

Private Function Rslt_2(A() As SwLnx) As SwErLy
Dim Sw As New Dictionary
Dim Ay() As SwLnx
    Dim M As Evl
    Dim J%
    Dim HasEvl As Boolean
    HasEvl = True
    Ay = A
    While HasEvl
        J = J + 1: If J > 1000 Then Stop
        M = Rslt_3(Ay, Sw)
        Set Sw = M.Sw
        Ay = M.Ay
        HasEvl = M.HasEvl
    Wend
With Rslt_2
    .ErLy = Rslt_4(Ay)
    Set .Sw = Rslt_5(Sw)
End With
End Function

Private Function Rslt_3(A() As SwLnx, Sw As Dictionary) As Evl
Dim OAy() As SwLnx
Dim OHasEvl As Boolean
Dim OSw As New Dictionary
Stop
'    Set OSw = DicClone(Sw)
    Dim I%, J%, SomLinEvaluated As Boolean
    SomLinEvaluated = True
    While SomLinEvaluated
        SomLinEvaluated = False
        I = I + 1
        If I > 1000 Then Stop
        For J = 0 To UB(A)
            With A(J).SomBool(OSw)
                If .Som Then
                    OHasEvl = True                  '<==
                    SomLinEvaluated = True
                    OSw.Add A(J).Nm, .Bool          '<==
                Else
                    PushObj OAy, A(J)               '<==
                End If
            End With
        Next
    Wend
With Rslt_3
    .Ay = OAy
    Set .Sw = OSw
    .HasEvl = OHasEvl
End With
End Function

Private Function Rslt_4(LeftOvr() As SwLnx) As String()
'Return ErLy for each of the LeftOvr
Stop
End Function

Private Function Rslt_5(A As Dictionary) As Sw
Stop
Set Rslt_5 = New Sw
End Function

Private Property Get U&()
U = N - 1
End Property

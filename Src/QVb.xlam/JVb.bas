Attribute VB_Name = "JVb"
Option Explicit

Property Get BoolAyOpt(A() As Boolean) As BoolAyOpt
Dim O As New BoolAyOpt
Set BoolAyOpt = O.Init(A)
End Property

Property Get BoolOpt(Bool As Boolean) As BoolOpt
Dim O As New BoolOpt
Set BoolOpt = O.Init(Bool)
End Property

Property Get DicOpt(A As Dictionary) As DicOpt
Dim O As New DicOpt
Set O.Dic = A
Set DicOpt = O
End Property

Property Get FmTo(Fmix&, Toix&) As FmTo
Dim O As New FmTo
Set FmTo = O.Init(Fmix, Toix)
End Property

Property Get FmToPos(FmPos, ToPos) As FmToPos
Dim O As New FmToPos
Set FmToPos = O.Init(FmPos, ToPos)
End Property

Property Get IntAyObj(Ay%()) As IntAyObj
Dim O As New IntAyObj
Set IntAyObj = O.Init(Ay)
End Property

Property Get IntOpt(I%) As IntOpt
Dim O As New IntOpt
Set IntOpt = O.Init(I)
End Property

Property Get KeyVal(K, V) As KeyVal
Dim O As New KeyVal
Set KeyVal = O.Init(K, V)
End Property

Property Get KeyValOpt(A As KeyVal) As KeyValOpt
Dim O As New KeyValOpt
Set KeyValOpt = O.Init(A)
End Property

Property Get LnoCnt(Lno&, Cnt&) As LnoCnt
Dim O As New LnoCnt
Set LnoCnt = O.Init(Lno, Cnt)
End Property

Property Get P123(P1, P2, P3) As P123
Dim O As New P123
Set P123 = O.Init(P1, P2, P3)
End Property

Property Get RRCC(R1&, R2&, C1&, C2&) As RRCC
Dim O As New RRCC
With O
    .R2 = R2
    .R1 = R1
    .C2 = C2
    .C1 = C1
End With
End Property

Property Get S1S2(S1, S2) As S1S2
Dim O As New S1S2
Set S1S2 = O.Init(S1, S2)
End Property

Property Get S1S2Opt(A As S1S2) As S1S2Opt
Dim O As New S1S2Opt
Set S1S2Opt = O.Init(A)
End Property

Property Get StrObj(A) As StrObj
Dim O As New StrObj
Set StrObj = O.Init(A)
End Property

'========================================================
Property Get StrOpt(S) As StrOpt
Dim O As New StrOpt
Set StrOpt = O.Init(S)
End Property

Property Get StrRslt(S, ErLy$()) As StrRslt
Dim O As New StrRslt
Set StrRslt = O.Init(S, ErLy)
End Property

Property Get SyObj(Sy$()) As SyObj
Dim O As New SyObj
Set SyObj = O.Init(Sy)
End Property

Property Get SyOpt(Sy$()) As SyOpt
Dim O As New SyOpt
Set SyOpt = O.Init(Sy)
End Property

Property Get SyPair(Sy1$(), Sy2$()) As SyPair
Dim O As New SyPair
Set SyPair = O.Init(Sy1, Sy2)
End Property

Property Get Tst() As Tst
Dim O As New Tst
Set Tst = O
End Property

Property Get ValOpt(V) As ValOpt
Dim O As New ValOpt
Set ValOpt = O.Init(V)
End Property

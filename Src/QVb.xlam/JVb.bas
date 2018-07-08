Attribute VB_Name = "JVb"
Option Explicit

Property Get ABC(Lin) As ABC
Dim O As New ABC
Set ABC = O.Init(Lin)
End Property

Property Get DDLines(Ly$()) As DDLines
Dim O As New DDLines
Set DDLines = O.Init(Ly)
End Property

Property Get Emp() As Emp
Static Y As New Emp
Set Emp = Y
End Property

Property Get FmTo(FmIx&, ToIx&) As FmTo
Dim O As New FmTo
Set FmTo = O.Init(FmIx, ToIx)
End Property

Property Get FmToPos(FmPos, ToPos) As FmToPos
Dim O As New FmToPos
Set FmToPos = O.Init(FmPos, ToPos)
End Property

Property Get Gp(A() As Lnx) As Gp
Dim O As New Gp
Set Gp = O.Init(A)
End Property

Property Get IntAyObj(Ay%()) As IntAyObj
Dim O As New IntAyObj
Set IntAyObj = O.Init(Ay)
End Property

Property Get KeyVal(K, V) As KeyVal
Dim O As New KeyVal
Set KeyVal = O.Init(K, V)
End Property

Property Get LABCAyRslt(A() As LABC, ErLy$()) As LABCAyRslt
Dim O As New LABCAyRslt
Set LABCAyRslt = O.Init(A, ErLy)
End Property

Property Get Lnx(Lin, Lx%) As Lnx
Dim O As New Lnx
Set Lnx = O.Init(Lin, Lx)
End Property

Property Get LyRslt(Ly$(), ErLy$()) As LyRslt
Dim O As New LyRslt
Set LyRslt = O.Init(Ly, ErLy)
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
Set RRCC = O
End Property

Property Get S1S2(S1, S2) As S1S2
Dim O As New S1S2
Set S1S2 = O.Init(S1, S2)
End Property

Property Get SomBool(Bool As Boolean) As SomBool
Dim O As New SomBool
Set SomBool = O.Init(Bool)
End Property

Property Get SomBoolAy(A() As Boolean) As SomBoolAy
Dim O As New SomBoolAy
Set SomBoolAy = O.Init(A)
End Property

Property Get SomDic(A As Dictionary) As SomDic
Dim O As New SomDic
Set O.Dic = A
Set SomDic = O
End Property

Property Get SomInt(I%) As SomInt
Dim O As New SomInt
Set SomInt = O.Init(I)
End Property

Property Get SomKeyVal(A As KeyVal) As SomKeyVal
Dim O As New SomKeyVal
Set SomKeyVal = O.Init(A)
End Property

Property Get SomS1S2(A As S1S2) As SomS1S2
Dim O As New SomS1S2
Set SomS1S2 = O.Init(A)
End Property

'========================================================
Property Get SomStr(S) As SomStr
Dim O As New SomStr
Set SomStr = O.Init(S)
End Property

Property Get SomSy(Sy$()) As SomSy
Dim O As New SomSy
Set SomSy = O.Init(Sy)
End Property

Property Get SomV(V) As SomV
Dim O As New SomV
Set SomV = O.Init(V)
End Property

Property Get StrObj(A) As StrObj
Dim O As New StrObj
Set StrObj = O.Init(A)
End Property

Property Get StrRslt(S, ErLy$()) As StrRslt
Dim O As New StrRslt
Set StrRslt = O.Init(S, ErLy)
End Property

Property Get SyObj(Sy$()) As SyObj
Dim O As New SyObj
Set SyObj = O.Init(Sy)
End Property

Property Get SyPair(Sy1$(), Sy2$()) As SyPair
Dim O As New SyPair
Set SyPair = O.Init(Sy1, Sy2)
End Property

Property Get Tst() As Tst
Dim O As New Tst
Set Tst = O
End Property

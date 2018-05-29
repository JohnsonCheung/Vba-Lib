Attribute VB_Name = "VbDic"
Option Explicit
Type KeyVal
   K As String
   V As Variant
End Type
Enum e_LinesLyOpt
    e_EscFstSpc = 1
End Enum
Enum e_DicLyOpt
    e_DrsFmt = 1
    e_KeyLinesFmt = 2
End Enum
Type KeyValOpt
   Som As Boolean
   KeyVal As KeyVal
End Type

Sub AssDicHasKeyLvs(A As Dictionary, KeyLvs$)
AssDicHasKy A, LvsSy(KeyLvs)
End Sub

Sub AssDicHasKy(A As Dictionary, Ky)
Dim K
For Each K In Ky
   If Not A.Exists(K) Then Debug.Print K: Stop
Next
End Sub

Sub AssEqDic(D1 As Dictionary, D2 As Dictionary)
If Not IsEqDic(D1, D2) Then Stop
End Sub

Function AyPair_Dic(A1, A2) As Dictionary
Dim N1&, N2&
N1 = Sz(A1)
N2 = Sz(A2)
If N1 <> N2 Then Stop
Dim O As New Dictionary
Dim J&
If AyIsEmp(A1) Then GoTo X
For J = 0 To N1 - 1
    O.Add A1(J), A2(J)
Next
X:
Set AyPair_Dic = O
End Function

Function DCRsltIsSam(A As DCRslt) As Boolean
With A
If .ADif.Count > 0 Then Exit Function
If .BDif.Count > 0 Then Exit Function
If .AExcess.Count > 0 Then Exit Function
If .BExcess.Count > 0 Then Exit Function
End With
DCRsltIsSam = True
End Function

Function DCRsltLy(A As DCRslt, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd")
With A
Dim A1() As S1S2: A1 = DCRsltS1S2Ay_Of_AExcess(.AExcess)
Dim A2() As S1S2: A2 = DCRsltS1S2Ay_Of_BExcess(.BExcess)
Dim A3() As S1S2: A3 = DCRsltS1S2Ay_Of_Dif(.ADif, .BDif)
Dim A4() As S1S2: A4 = DCRsltS1S2Ay_Of_Sam(.Sam)
End With
Dim O() As S1S2
S1S2_Push O, NewS1S2(Nm1, Nm2)
O = S1S2_Add(O, A1)
O = S1S2_Add(O, A2)
O = S1S2_Add(O, A3)
O = S1S2_Add(O, A4)
DCRsltLy = S1S2Ay_FmtLy(O)
End Function

Function DCRsltS1S2Ay_Of_AExcess(AExcess As Dictionary) As S1S2()
If DicIsEmp(AExcess) Then Exit Function
Dim O() As S1S2, K
For Each K In AExcess.Keys
    S1S2_Push O, NewS1S2(K & vbCrLf & StrDup(Len(K), "-") & vbCrLf & AExcess(K), "")
Next
DCRsltS1S2Ay_Of_AExcess = O
End Function

Function DCRsltS1S2Ay_Of_BExcess(BExcess As Dictionary) As S1S2()
If DicIsEmp(BExcess) Then Exit Function
Dim O() As S1S2, K
For Each K In BExcess.Keys
    S1S2_Push O, NewS1S2("", K & vbCrLf & StrDup(Len(K), "-") & vbCrLf & BExcess(K))
Next
DCRsltS1S2Ay_Of_BExcess = O
End Function

Function DCRsltS1S2Ay_Of_Dif(ADif As Dictionary, BDif As Dictionary) As S1S2()
If DicSz(ADif) <> DicSz(BDif) Then Stop
If DicIsEmp(ADif) Then Exit Function
Dim O() As S1S2, K, S1$, S2$
For Each K In ADif
    S1 = K & vbCrLf & StrDup(Len(K), "-") & vbCrLf & ADif(K)
    S2 = K & vbCrLf & StrDup(Len(K), "-") & vbCrLf & BDif(K)
    S1S2_Push O, NewS1S2(S1, S2)
Next
DCRsltS1S2Ay_Of_Dif = O
End Function

Function DCRsltS1S2Ay_Of_Sam(ASam As Dictionary) As S1S2()
If DicIsEmp(ASam) Then Exit Function
Dim O() As S1S2, K
For Each K In ASam.Keys
    S1S2_Push O, NewS1S2("*Same", K & vbCrLf & StrDup(Len(K), "-") & vbCrLf & ASam(K))
Next
DCRsltS1S2Ay_Of_Sam = O
End Function

Function DicAdd(A As Dictionary, ParamArray DicAp()) As Dictionary
Dim Av(): Av = DicAp
Dim I, Dic As Dictionary
Dim O As Dictionary
Set O = DicClone(A)
For Each I In Av
   Set Dic = I
   Set O = DicAddOne(O, Dic)
Next
Set DicAdd = O
End Function

Function DicAddKeyPfx(A As Dictionary, Pfx) As Dictionary
Dim O As New Dictionary
If A.Count = 0 Then GoTo X
Dim K
For Each K In A.Keys
   O.Add Pfx & K, A(K)
Next
X:
   Set DicAddKeyPfx = O
End Function

Sub DicAddKeyVal(A As Dictionary, KeyVal As KeyVal)
With KeyVal
   A.Add .K, .V
End With
End Sub

Sub DicAddKeyValOpt(A As Dictionary, KeyValOpt As KeyValOpt)
With KeyValOpt
   If .Som Then DicAddKeyVal A, .KeyVal
End With
End Sub

Function DicAddOne(A As Dictionary, B As Dictionary) As Dictionary
Dim O As Dictionary: Set O = DicClone(A)
Dim K
If B.Count > 0 Then
   For Each K In B.Keys
       O.Add K, B(K)
   Next
End If
Set DicAddOne = O
End Function

Function DicAyAdd(Dy() As Dictionary) As Dictionary
Dim O As Dictionary
   Set O = DicClone(Dy(0))
Dim J%
For J = 1 To UB(Dy)
   Set O = DicAddOne(O, Dy(J))
Next
Set DicAyAdd = O
End Function

Function DicAyDr(DicAy, K) As Variant()
Dim U%: U = UB(DicAy)
Dim O()
ReDim O(U + 1)
Dim I, Dic As Dictionary, J%
J = 1
O(0) = K
For Each I In DicAy
   Set Dic = I
   If Dic.Exists(K) Then O(J) = Dic(K)
   J = J + 1
Next
DicAyDr = O
End Function

Function DicAyKy(DicAy) As Variant()
Dim O(), Dic As Dictionary, I
For Each I In DicAy
   Set Dic = I
   PushNoDupAy O, Dic.Keys
Next
DicAyKy = O
End Function

Function DicBoolOpt(A As Dictionary, K) As BoolOpt
Dim V As VOpt: V = DicValOpt(A, K)
If V.Som Then DicBoolOpt = SomBool(V.V)
End Function

Sub DicBrw(A As Dictionary)
DrsBrw DicDrs(A)
End Sub

Function DicByDry(DicDry) As Dictionary
Dim O As New Dictionary
If Not AyIsEmp(DicDry) Then
   Dim Dr
   For Each Dr In DicDry
       O.Add Dr(0), Dr(1)
   Next
End If
Set DicByDry = O
End Function

Function DicClone(A As Dictionary) As Dictionary
Dim O As New Dictionary, K
If A.Count > 0 Then
   For Each K In A.Keys
       O.Add K, A(K)
   Next
End If
Set DicClone = O
End Function

Sub DicCmp(A As Dictionary, B As Dictionary)
AyBrw DCRsltLy(DicCmpRslt(A, B))
End Sub

Function DicCmpRslt(A As Dictionary, B As Dictionary, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd") As DCRslt
Dim O As DCRslt
Set O.AExcess = DicMinus(A, B)
Set O.BExcess = DicMinus(B, A)
Set O.Sam = DicIntersect(A, B)
With DicSamKeyDifValPair(A, B)
    Set O.ADif = .A
    Set O.BDif = .B
End With
O.Nm1 = Nm1
O.Nm2 = Nm2
DicCmpRslt = O
End Function

Sub DicDmp(A As Dictionary, Optional InclDicValTy As Boolean, Optional Opt As e_DicLyOpt = e_DrsFmt)
AyDmp DicLy(A, InclDicValTy, Opt)
End Sub

Function DicDrs(A As Dictionary, Optional InclDicValTy As Boolean) As Drs
Dim O As Drs
O.Fny = SplitSpc("Key Val"): If InclDicValTy Then Push O.Fny, "ValTy"
O.Dry = DicDry(A, InclDicValTy)
DicDrs = O
End Function

Function DicDry(A As Dictionary, Optional InclDicValTy As Boolean) As Variant()
Dim O(), I
If A.Count = 0 Then Exit Function
Dim K(): K = A.Keys
If Not AyIsEmp(K) Then
   If InclDicValTy Then
       For Each I In K
           Push O, Array(I, A(I), TypeName(A(I)))
       Next
   Else
       For Each I In K
           Push O, Array(I, A(I))
       Next
   End If
End If
DicDry = O
End Function

Function DicDt(A As Dictionary, Optional DtNm$ = "Dic", Optional InclDicValTy As Boolean) As Dt
DicDt = NewDtByDrs(DtNm, DicDrs(A, InclDicValTy))
End Function

Function DicHasBlankKey(A As Dictionary) As Boolean
If DicIsEmp(A) Then Exit Function
Dim K
For Each K In A.Keys
   If Trim(K) = "" Then DicHasBlankKey = True: Exit Function
Next
End Function

Function DicHasK(A As Dictionary, K$) As Boolean
DicHasK = A.Exists(K)
End Function

Function DicHasKeyLvs(A As Dictionary, KeyLvs) As Boolean
DicHasKeyLvs = DicHasKy(A, LvsSy(KeyLvs))
End Function

Function DicHasKy(A As Dictionary, Ky) As Boolean
Ass IsArray(Ky)
If AyIsEmp(Ky) Then Stop
Dim K
For Each K In Ky
   If Not A.Exists(K) Then
       Debug.Print FmtQQ("DicHasKy: Key(?) is missing", K)
       Exit Function
   End If
Next
DicHasKy = True
End Function

Function DicIntersect(A As Dictionary, B As Dictionary) As Dictionary
Dim O As New Dictionary
If DicIsEmp(A) Then GoTo X
If DicIsEmp(B) Then GoTo X
Dim K
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) = B(K) Then
            O.Add K, A(K)
        End If
    End If
Next
X: Set DicIntersect = O
End Function

Function DicIsEmp(A As Dictionary) As Boolean
If IsNothing(A) Then DicIsEmp = True: Exit Function
DicIsEmp = A.Count = 0
End Function

Function DicJn(DicAy, Optional FnyOpt) As Drs
Const CSub$ = "DicJn"
Dim UDic%
   UDic = UB(DicAy)
Dim Fny$()
   If VarIsEmp(FnyOpt) Then
       Dim J%
       Push Fny, "Key"
       For J = 0 To UDic
           Push Fny, "V" & J
       Next
   Else
       Fny = FnyOpt
   End If
If UB(Fny) <> UDic + 1 Then Er CSub, "Given {FnyOpt} has {Sz} <> {DicAy-Sz}", FnyOpt, Sz(FnyOpt), Sz(DicAy)
Dim Ky()
   Ky = DicAyKy(DicAy)
Dim URow&
   URow = UB(Ky)
Dim O()
   ReDim O(URow)
   Dim K
   J = 0
   For Each K In Ky
       O(J) = DicAyDr(DicAy, K)
       J = J + 1
   Next
DicJn.Dry = O
DicJn.Fny = Fny
End Function

Function DicKVLy(A As Dictionary) As String()
If DicIsEmp(A) Then Exit Function
Dim O$(), K, W%, Ky
Ky = A.Keys
W = AyWdt(Ky)
For Each K In Ky
   Push O, AlignL(K, W) & " = " & A(K)
Next
DicKVLy = O
End Function

Function DicKeySy(A As Dictionary) As String()
DicKeySy = AySy(A.Keys)
End Function

Function DicLines_Dic(A$) As Dictionary
Set DicLines_Dic = DicLy_Dic(SplitLines(A))
End Function

Function DicLy(A As Dictionary, Optional InclDicValTy As Boolean, Optional Opt As e_DicLyOpt = e_DrsFmt) As String()
Select Case Opt
Case e_DrsFmt: DicLy = DrsLy(DicDrs(A, InclDicValTy))
Case e_KeyLinesFmt: DicLy = S1S2Ay_KeyLinesLy(DicS1S2Ay(A))
Case Else: Stop
End Select
End Function

Function DicLy_Dic(A$(), Optional JnSep$ = vbCrLf) As Dictionary
Const CSub$ = "DicLy_Dic"
Dim O As New Dictionary
   If AyIsEmp(A) Then Set DicLy_Dic = O: Exit Function
   Dim I
   For Each I In A
       If Trim(I) = "" Then GoTo Nxt
       If FstChr(I) = "#" Then GoTo Nxt
       With Brk(I, " ")
           If O.Exists(.S1) Then
               O(.S1) = O(.S1) & JnSep & .S2
           Else
               O.Add .S1, .S2
           End If
       End With
Nxt:
   Next
Set DicLy_Dic = O
End Function

Function DicMge(PfxLvs$, ParamArray DicAp()) As Dictionary
Dim Av(): Av = DicAp
Dim Ny$()
   Ny = LvsSy(PfxLvs)
   Ny = AyAddSfx(Ny, "@")
If Sz(Av) <> Sz(Ny) Then Stop
Dim Dy() As Dictionary
Dim D As Dictionary
   Dim J%
   For J = 0 To UB(Ny)
       Set D = Av(J)
       Push Dy, DicAddKeyPfx(D, Ny(J))
   Next
Set DicMge = DicAyAdd(Dy)
End Function

Function DicMinus(A As Dictionary, B As Dictionary) As Dictionary
If DicIsEmp(A) Then Set DicMinus = New Dictionary: Exit Function
If DicIsEmp(B) Then Set DicMinus = DicClone(A): Exit Function
Dim O As New Dictionary, K
For Each K In A.Keys
   If Not B.Exists(K) Then O.Add K, A(K)
Next
Set DicMinus = O
End Function

Function DicS1S2Ay(A As Dictionary) As S1S2()
If DicIsEmp(A) Then Exit Function
Dim O() As S1S2
Dim U&: U = A.Count - 1
ReDim O(U)
Dim J&, I
For Each I In A
    O(J).S1 = I
    O(J).S2 = A(I)
    J = J + 1
Next
DicS1S2Ay = O
End Function

Function DicSamKeyDifValPair(A As Dictionary, B As Dictionary) As DicPair
Dim K, A1 As New Dictionary, B1 As New Dictionary
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) <> B(K) Then
            A1.Add K, A(K)
            B1.Add K, B(K)
        End If
    End If
Next
With DicSamKeyDifValPair
    Set .A = A1
    Set .B = B1
End With
End Function

Function DicSelIntoAy(A As Dictionary, Ky$()) As Variant()
Dim O()
Dim U&: U = UB(Ky)
ReDim O(U)
Dim J&
For J = 0 To U
   If Not A.Exists(Ky(J)) Then Stop
   O(J) = A(Ky(J))
Next
DicSelIntoAy = O
End Function

Function DicSelIntoSy(A As Dictionary, Ky$()) As String()
DicSelIntoSy = AySy(DicSelIntoAy(A, Ky))
End Function

Function DicSrt(A As Dictionary) As Dictionary
If DicIsEmp(A) Then Set DicSrt = New Dictionary: Exit Function
Dim K
Dim O As New Dictionary
For Each K In AySrt(A.Keys)
   O.Add K, A(K)
Next
Set DicSrt = O
End Function

Function DicStrKy(A As Dictionary) As String()
DicStrKy = AySy(A.Keys)
End Function

Function DicSz&(A As Dictionary)
If IsNothing(A) Then Exit Function
DicSz = A.Count
End Function

Function DicToLy(A As Dictionary) As String()
If DicIsEmp(A) Then Exit Function
Dim Key: Key = A.Keys
Dim O$(): O = AyAlignL(Key)
Dim J&
For J = 0 To UB(Key)
   O(J) = O(J) & " " & A(Key(J))
Next
DicToLy = O
End Function

Function DicVal(A As Dictionary, K, Optional ThrowErIfNotFnd As Boolean)
If Not A.Exists(K) Then
   If ThrowErIfNotFnd Then Stop
   Exit Function
End If
DicVal = A(K)
End Function

Function DicValOpt(A As Dictionary, K) As VOpt
If Not A.Exists(K) Then Exit Function
DicValOpt = SomV(A(K))
End Function

Function DicVbl_Dic(A$, Optional JnSep$ = vbCrLf) As Dictionary
Set DicVbl_Dic = DicLy_Dic(SplitVBar(A), JnSep)
End Function

Function DicWs(A As Dictionary) As Worksheet
Set DicWs = DrsWs(DicDrs(A))
End Function

Function DicWsVis(A As Dictionary) As Worksheet
Dim O As Worksheet
   Set O = DicWs(A)
   WsVis O
Set DicWsVis = O
End Function

Function FtDic(FT) As Dictionary
Set FtDic = DicLy_Dic(FtLy(FT))
End Function

Function IsEqDic(D1 As Dictionary, D2 As Dictionary) As Boolean
If DicIsEmp(D1) Then Stop
If DicIsEmp(D2) Then Stop
If D1.Count <> D2.Count Then Exit Function
Dim K1, K2
K1 = AySrt(D1.Keys)
K2 = AySrt(D2.Keys)
If AyIsEq(K1, K2) Then Exit Function
Dim K
For Each K In K1
   If D1(K) <> D2(K) Then Exit Function
Next
IsEqDic = True
End Function

Function IsVdtLyDicStr(A) As Boolean
If Left(A, 3) <> "***" Then Exit Function
Dim I, K$(), Key$
For Each I In SplitCrLf(A)
   If Left(I, 3) = "***" Then
       Key = Mid(I, 4)
       If AyHas(K, Key) Then Exit Function
       Push K, Key
   End If
Next
IsVdtLyDicStr = True
End Function

Function KeyLines_Ly(K, Lines, Optional KeyWdt0%) As String()
Dim Ly$(): Ly = LinesLy(Lines, e_EscFstSpc)
Dim W%: If Len(K) > KeyWdt0 Then W = Len(K) Else W = KeyWdt0
KeyLines_Ly = AyAddPfx(Ly, AlignL(K, W) & " ")
End Function

Function KeyVal(K$, V) As KeyVal
KeyVal.K = K
KeyVal.V = V
End Function

Function LinesDicLy_LinesDic(A$()) As Dictionary
Dim T1Ay$(), RstAy$()
    With LyT1AyRstAy(A)
        T1Ay = .T1Ay
        RstAy = .RstAy
    End With
Dim Ny$()
    Ny = AyNoDupAy(T1Ay)
Dim O As Dictionary
    Dim Lines$, J&
    Set O = New Dictionary
    For J = 0 To UB(Ny)
        Lines = LinesDicLy_LinesDic__Lines(T1Ay, RstAy, Ny(J))
        O.Add Ny(J), Lines
    Next
Set LinesDicLy_LinesDic = O
End Function

Sub LinesDic_Brw(A As Dictionary)
AyBrw S1S2Ay_FmtLy(LinesDic_S1S2Ay(A))
End Sub

Function LinesDic_Ly(A As Dictionary) As String()
Dim O$(), K
If DicIsEmp(A) Then Exit Function
For Each K In A.Keys
    Push O, LinesDic_Ly__Ly(K, A(K))
Next
LinesDic_Ly = O
End Function

Function LinesDic_Ly__Ly(K, Lines) As String()
Dim O$(), J&
Dim Ly$()
    Ly = SplitCrLf(Lines)
For J = 0 To UB(Ly)
    Dim Lin$
        Lin = Ly(J)
        If FstChr(Lin) = " " Then Lin = "~" & RmvFstChr(Lin)
    Push O, K & " " & Lin
Next
LinesDic_Ly__Ly = O
End Function

Function LinesDic_S1S2Ay(A As Dictionary) As S1S2()
If DicIsEmp(A) Then Exit Function
Dim O() As S1S2
ReDim O(A.Count - 1)
Dim J&, K
For Each K In A.Keys
    O(J) = NewS1S2(K, A(K))
    J = J + 1
Next
LinesDic_S1S2Ay = O
End Function

Function LinesLy(Lines, Optional Opt As e_LinesLyOpt = e_EscFstSpc) As String()
Dim L$(), J&
L = SplitCrLf(Lines)
Select Case Opt
Case e_EscFstSpc
    For J = 0 To UB(L)
        If FstChr(L(J)) = " " Then L(J) = RplFstChr(L(J), "~")
    Next
End Select
LinesLy = L
End Function

Function LyBoolDic(A$()) As Dictionary
Dim Z As New Dictionary, J%
For J = 0 To UB(A)
    With Brk1(A(J), " ")
        Z.Add .S1, CBool(.S2)
    End With
Next
Set LyBoolDic = Z
End Function

Function LyDic_FmtLines(A As Dictionary) As String
LyDic_FmtLines = JnCrLf(LyDic_FmtLy(A))
End Function

Function LyDic_FmtLy(A As Dictionary) As String()
Dim O$()
If DicIsEmp(A) Then LyDic_FmtLy = ApSy("***"): Exit Function
Dim K
For Each K In A.Keys
   Push O, "***" & K
   PushAy O, NewLyDic(K)
Next
LyDic_FmtLy = O
End Function

Function LyDic_Wb(A As Dictionary, Optional Vis As Boolean) As Workbook
'LyDic is a dictionary with K is string and V is Ly
Dim O As Workbook: Set O = NewWb
If DicIsEmp(A) Then GoTo X
Dim Ws As Worksheet, K, ThereIsSheet1 As Boolean
For Each K In A.Keys
    If K = "Sheet1" Then
        Set Ws = O.Sheets("Sheet1")
        ThereIsSheet1 = True
    Else
        Set Ws = O.Sheets.Add
        Ws.Name = K
    End If
    AyRgV(A(K), WsA1(Ws)).Font.Name = "Courier New"
Next
If Not ThereIsSheet1 Then
    WbWs(O, "Sheet1").Delete
End If
If Vis Then O.Application.Visible = True
X: Set LyDic_Wb = O
End Function

Function NewDic() As Dictionary
Set NewDic = New Dictionary
End Function

Function NewDicByLysDicStr(LyDicStr$) As Dictionary
'SpecStr:LyDicStr.  It is LyDic-Str.  It is a str which can made a LinesDic.  Of format ***Key1|Lines1....|***Key2|Lines2....|
Ass IsVdtLyDicStr(LyDicStr)
Dim A$(): A = Split(LyDicStr, "***")
If A(0) <> "" Then Stop
A = AyRmvEleAt(A)
Dim A1()
    Dim J%
    For J = 0 To UB(A)
        Push A1, SplitCrLf(A(J))
    Next
End Function

Function NewLyDic(LyDicStr) As Dictionary
Ass IsVdtLyDicStr(LyDicStr)
Dim A$(): A = Split(LyDicStr, "***")
Dim J%, O As New Dictionary
For J = 1 To UB(A)
    Dim B$()
    B = SplitCrLf(A(J))
    O.Add B(0), AyRmvEleAt(B)
Next
Set NewLyDic = O
End Function

Function S1S2Ay_Dic(A() As S1S2) As Dictionary
Dim J&, O As New Dictionary
For J = 0 To S1S2_UB(A)
    With A(J)
        O.Add .S1, .S2
    End With
Next
Set S1S2Ay_Dic = O
End Function

Function S1S2Ay_KeyLinesLy(A() As S1S2) As String()
Dim U%: U = S1S2_UB(A)
If U = -1 Then Exit Function
Dim O$(), W%, J&
W = S1S2Ay_Wdt1(A)
For J = 0 To U
    With A(J)
        PushAy O, KeyLines_Ly(.S1, .S2, W)
    End With
Next
S1S2Ay_KeyLinesLy = O
End Function

Function SomKeyVal(K$, V) As KeyValOpt
SomKeyVal.Som = True
SomKeyVal.KeyVal = KeyVal(K, V)
End Function

Sub ZZ_DicCmp()
Dim A As Dictionary: Set A = DicVbl_Dic("X AA|A BBB|A Lines1|A Line3|B Line1|B line2|B line3..")
Dim B As Dictionary: Set B = DicVbl_Dic("X AA|C Line|D Line1|D line2|B Line1|B line2|B line3|B Line4")
DicCmp A, B
End Sub

Private Function LinesDicLy_LinesDic__Lines$(FstTermAy$(), RstAy$(), Nm$)
Dim O$(), J&
For J = 0 To UB(FstTermAy)
    If FstTermAy(J) = Nm Then
        Dim Lin$
        Lin = RstAy(J)
        If FstChr(Lin) = "~" Then Lin = RplFstChr(Lin, " ")
        Push O, Lin
    End If
Next
LinesDicLy_LinesDic__Lines = JnCrLf(O)
End Function

Sub DicS1S2Ay__Tst()
Dim A As New Dictionary
A.Add "A", "BB"
A.Add "B", "CCC"
Dim Act() As S1S2
Act = DicS1S2Ay(A)
Stop
End Sub

Private Sub IsVdtLyDicStr__Tst()
Ass IsVdtLyDicStr(RplVBar("***ksdf|***ksdf1")) = True
Ass IsVdtLyDicStr(RplVBar("***ksdf|***ksdf")) = False
Ass IsVdtLyDicStr(RplVBar("**ksdf|***ksdf")) = False
Ass IsVdtLyDicStr(RplVBar("***")) = True
Ass IsVdtLyDicStr("**") = False
End Sub

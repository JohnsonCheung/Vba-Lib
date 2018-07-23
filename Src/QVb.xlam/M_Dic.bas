Attribute VB_Name = "M_Dic"
Option Explicit

Function DicAdd(A As Dictionary, B As Dictionary) As Dictionary
Dim O As Dictionary: Set O = DicClone(A)
Dim K
If B.Count > 0 Then
   For Each K In B.Keys
       O.Add K, B(K)
   Next
End If
Set DicAdd = O
End Function

Function DicAddAp(A As Dictionary, ParamArray DicAp()) As Dictionary
Dim Av(): Av = DicAp
Dim I, Dic As Dictionary
Dim O As Dictionary
Set O = DicClone(A)
For Each I In Av
   Set O = DicAdd(O, CvDic(I))
Next
Set DicAddAp = O
End Function

Function DicAddAy(A As Dictionary, Dy() As Dictionary) As Dictionary
Dim O As Dictionary
   Set O = DicClone(A)
Dim J%
For J = 0 To UB(Dy)
   Set O = DicAdd(O, Dy(J))
Next
Set DicAddAy = O
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

Function DicAllKeyIsNm(A As Dictionary) As Boolean
Dim K
For Each K In A.Keys
    If Not IsNm(K) Then Exit Function
Next
DicAllKeyIsNm = True
End Function

Function DicAllKeyIsStr(A As Dictionary) As Boolean
DicAllKeyIsStr = AyIsAllStr(A.Keys)
End Function

Function DicAllValIsStr(A As Dictionary) As Boolean
DicAllValIsStr = AyIsAllStr(A.Items)
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

Function DicClone(A As Dictionary) As Dictionary
Dim O As New Dictionary, K
If A.Count > 0 Then
   For Each K In A.Keys
       O.Add K, A(K)
   Next
End If
Set DicClone = O
End Function

Function DicDry(A As Dictionary, Optional InclValTy As Boolean) As Variant()
If A.Count = 0 Then Exit Function
Dim O(), K, V
For Each K In A.Keys
    If InclValTy Then
        Push O, Array(K, A(K))
    Else
        V = A(K)
        Push O, Array(K, V, TypeName(V))
    End If
Next
DicDry = O
End Function

Function DicDry_Dic(DicDry()) As Dictionary
Dim O As New Dictionary
If Sz(DicDry) > 0 Then
   Dim Dr
   For Each Dr In DicDry
       O.Add Dr(0), Dr(1)
   Next
End If
Set DicDry_Dic = O
End Function

Function DicFny(Optional InclValTy As Boolean) As String()
DicFny = SslSy("Key Val" & IIf(InclValTy, " Type", ""))
End Function

Function DicHasBlankKey(A As Dictionary) As Boolean
If A.Count = 0 Then Exit Function
Dim K
For Each K In A.Keys
   If Trim(K) = "" Then DicHasBlankKey = True: Exit Function
Next
End Function

Function DicHasKeySsl(A As Dictionary, KeySsl) As Boolean
DicHasKeySsl = A.Exists(SslSy(KeySsl))
End Function

Function DicHasKy(A As Dictionary, Ky) As Boolean
Ass IsArray(Ky)
If AyIsEmp(Ky) Then Stop
Dim K
For Each K In Ky
   If Not A.Exists(K) Then
       Debug.Print FmtQQ("Dix.HasKy: Key(?) is missing", K)
       Exit Function
   End If
Next
DicHasKy = True
End Function

Function DicIntersect(A As Dictionary, B As Dictionary) As Dictionary
Dim O As New Dictionary
If A.Count = 0 Then GoTo X
If B.Count = 0 Then GoTo X
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

Function DicIsEq(A As Dictionary, B As Dictionary) As Boolean
If A.Count = 0 Then
    If B.Count = 0 Then
        DicIsEq = True
        Exit Function
    End If
    Exit Function
End If
If A.Count = 0 Then Exit Function
If A.Count <> B.Count Then Exit Function
Dim K1, K2
K1 = AySrt(A.Keys)
K2 = AySrt(B.Keys)
If AyIsEq(K1, K2) Then Exit Function
Dim K
For Each K In K1
   If B(K) <> A(K) Then Exit Function
Next
DicIsEq = True
End Function

Function DicKeySy(A As Dictionary) As String()
DicKeySy = AySy(A.Keys)
End Function

Function DicLines(A As Dictionary) As String
DicLines = JnCrLf(DicLy(A))
End Function

Function DicLines_Dic(DicLines, Optional JnSep$ = vbCrLf) As Dictionary
Set DicLines_Dic = DicLy_Dic(SplitLines(DicLines), JnSep)
End Function

Function DicLy(A As Dictionary, Optional InclDicValTy As Boolean) As String()
DicLy = S1S2Ay_Ly(DicS1S2Ay(A), IsAlignS1:=True)
End Function

Function DicLy1(A As Dictionary) As String()
If A.Count = 0 Then Exit Function
Dim Key: Key = A.Keys
Dim O$(): O = AyAlignL(Key)
Dim J&
For J = 0 To UB(Key)
   O(J) = O(J) & " " & A(Key(J))
Next
DicLy1 = O
End Function

Function DicLy2(A As Dictionary) As String()
Dim O$(), K
If A.Count = 0 Then Exit Function
For Each K In A.Keys
    Push O, DicLy2__1(K, A(K))
Next
DicLy2 = O
End Function

Function DicLy2__1(K, Lines) As String()
Dim O$(), J&
Dim Ly$()
    Ly = SplitCrLf(Lines)
For J = 0 To UB(Ly)
    Dim Lin$
        Lin = Ly(J)
        If FstChr(Lin) = " " Then Lin = "~" & RmvFstChr(Lin)
    Push O, K & " " & Lin
Next
DicLy2__1 = O
End Function

Function DicLy_Dic(A$(), Optional JnSep$ = vbCrLf) As Dictionary
Dim O As New Dictionary
Dim A1$(): A1 = AyRmvEmpEleAtEnd(A)
If AyIsEmp(A) Then Set DicLy_Dic = O: Exit Function
Dim I, T1$, Rst$
For Each I In A
    With LinT1Rst(I)
        T1 = .T1
        Rst = .Rst
    End With
    If O.Exists(T1) Then
        If FstChr(Rst) = "~" Then Rst = RplFstChr(Rst, " ")
        O(T1) = O(T1) & JnSep & Rst
    Else
        O.Add T1, Rst
    End If
 Next
Set DicLy_Dic = O
End Function

Function DicMge(A As Dictionary, PfxSsl$, ParamArray DicAp()) As Dictionary
Dim Av(): Av = DicAp
Dim Ny$()
   Ny = SslSy(PfxSsl)
   Ny = AyAddSfx(Ny, "@")
If Sz(Av) <> Sz(Ny) Then Stop
Dim Dy() As Dictionary
Dim D As Dictionary
   Dim J%
   For J = 0 To UB(Ny)
       Set D = Av(J)
       Push Dy, DicAddKeyPfx(A, Ny(J))
   Next
Set DicMge = DicAddAy(A, Dy)
End Function

Function DicMinus(A As Dictionary, B As Dictionary) As Dictionary
If A.Count = 0 Then Set DicMinus = New Dictionary: Exit Function
If B.Count = 0 Then Set DicMinus = DicClone(A): Exit Function
Dim O As New Dictionary, K
For Each K In A.Keys
   If Not B.Exists(K) Then O.Add K, A(K)
Next
Set DicMinus = O
End Function

Function DicS1S2Ay(A As Dictionary) As S1S2()
If A.Count = 0 Then Exit Function
Dim O() As S1S2
ReDim O(A.Count - 1)
Dim J&, K
For Each K In A.Keys
    Set O(J) = S1S2(K, A(K))
    J = J + 1
Next
DicS1S2Ay = O
End Function

Function DicSrt(A As Dictionary) As Dictionary
If A.Count = 0 Then Set DicSrt = New Dictionary: Exit Function
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

Function DicToStr$(A)
Dim O$(), K
For Each K In A.Keys
    Push O, KeyVal(K, A(K)).ToStr
Next
DicToStr = Tag("Dic", JnCrLf(O))
End Function

Function DicValOpt(A As Dictionary, K) As ValOpt
If Not A.Exists(K) Then Set DicValOpt = New ValOpt: Exit Function
Set DicValOpt = JVb.ValOpt(A(K))
End Function

Function DicWb(A As Dictionary, Optional Vis As Boolean) As Workbook
'Assume each dic keys is name and each value is lines
'Prp-Wb is to create a new Wb with worksheet as the dic key and the lines are break to each cell of the sheet
Ass DicAllKeyIsNm(A)
Ass DicAllValIsStr(A)
Dim K, ThereIsSheet1 As Boolean
Dim O As Workbook: Stop 'Set O = NewWb
Dim Ws As Worksheet
For Each K In A.Keys
    If K = "Sheet1" Then
        Set DicWb = O
        ThereIsSheet1 = True
    Else
        Set DicWs = O.Sheets.Add
        Ws.Name = K
    End If
    Ws.Range("A1").Value = LinesSqV(A(K))
Next
X: Set Ws = O
If Vis Then O.Application.Visible = True
End Function

Function DicWs(Optional Vis As Boolean) As Worksheet
Dim O As Worksheet: Set O = NewWs
Stop '
Set DicWs = O
End Function

Function Dic_WhEqKy(A As Dictionary, EqKy$()) As Variant()
Dim O()
Dim U&: U = UB(EqKy)
ReDim O(U)
Dim J&
For J = 0 To U
   If Not A.Exists(EqKy(J)) Then Stop ' All EqKy should exist in A
   Asg A(EqKy(J)), O(J)
Next
Dic_WhEqKy = O
End Function

Function Dic_WhEq_AsSy(A As Dictionary, Ky$()) As String()
Dic_WhEq_AsSy = AySy(Dic_WhEqKy(A, Ky))
End Function

Function S1S2Ay_Ly(A() As S1S2, Optional IsAlignS1 As Boolean) As String()

End Function

Sub DicBrw(A As Dictionary, Optional InclDicValTy As Boolean)
AyBrw DicLy(A, InclDicValTy)
End Sub

Sub DicDmp(A As Dictionary, Optional InclDicValTy As Boolean)
AyDmp DicLy(A, InclDicValTy)
End Sub

Sub DicPushKeyVal(A As Dictionary, KeyVal As KeyVal, Optional ThwEr As Boolean)
With KeyVal
    If A.Exists(.K) Then
        If ThwEr Then
            Er "DicPushKeyVal: Given {KeyVal.K} exists in {Dix}", KeyVal.K, DicToStr(A)
        Else
            Debug.Print "DicPushKeyVal: Given {KeyVal.K} exists in {Dix}.  Skip adding"
        End If
        Exit Sub
    End If
    A.Add .K, .V
End With
End Sub

Sub DicPushKeyValOpt(A As KeyValOpt)
With A
   If .Som Then DicPushKeyVal A, .KeyVal
End With
End Sub

Private Function DicLy__1(A As Dictionary) As String()
If A.Count = 0 Then Exit Function
Dim O$(), K, W%, Ky
Ky = A.Keys
W = AyWdt(Ky)
For Each K In Ky
   Push O, AlignL(K, W) & " " & A(K)
Next
DicLy__1 = O
End Function

Sub ZZ__Tst()
'ZZ_Cmp
'ZZ_S1S2s
ZZ_ToStr
End Sub

Private Sub ZZ_DicS1S2Ay()
Dim A As New Dictionary
A.Add "A", "BB"
A.Add "B", "CCC"
Dim Act() As S1S2
Act = DicS1S2Ay(A)
End Sub

Private Sub ZZ_ToStr()
Debug.Print VblDic("a b|c d|e x").ToStr
End Sub

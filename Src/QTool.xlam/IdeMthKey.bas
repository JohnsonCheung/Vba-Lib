Attribute VB_Name = "IdeMthKey"
Option Explicit





Sub Z_VbeMthDry()
Brw DryFmtss(VbeMthDry(CurVbe))
End Sub

Sub Z_PjMthDry()
Brw DryFmtss(PjMthDry(CurPj))
End Sub



Function ShfAs(A) As Variant()
Dim L$
L = LTrim(A)
If Left(L, 3) = "As " Then ShfAs = Array(True, LTrim(Mid(L, 4))): Exit Function
ShfAs = Array(False, A)
End Function

Function ShfTerm$(OLin$)
Dim L$, P%
L = LTrim(OLin)
If FstChr(L) = "[" Then
    P = SqBktEndPos(L)
    ShfTerm = Mid(L, 2, P - 2)
    OLin = LTrim(Mid(L, P + 1))
    Exit Function
End If
P = InStr(L, " ")
If P = 0 Then
    ShfTerm = L
    OLin = ""
    Exit Function
End If
ShfTerm = Left(L, P - 1)
OLin = Trim(Mid(L, P + 1))
End Function
Sub Z_VbeMthLinDry()
Brw DryFmtss(VbeMthLinDry(CurVbe))
End Sub

Function PjMthLinDry(A As VBProject) As Variant()
Dim M
For Each M In AyNz(PjMdAy(A))
    PushAy PjMthLinDry, MdMthLinDry(CvMd(M))
Next
End Function

Sub PushNonBlankAy(O, M)
If Sz(M) > 0 Then Push O, M
End Sub



Sub AAAA()
Z_VbeMthLinDryWP
End Sub

Sub Z_VbeMthLinDryWP()
Brw DryFmtssWrp(VbeMthLinDryWP(CurVbe))
End Sub


Function PjMthLinDryWP(A As VBProject) As Variant()
Dim M
For Each M In AyNz(PjMdAy(A))
    PushIAy PjMthLinDryWP, MdMthLinDryWP(CvMd(M))
Next
End Function




Function LinMthDrWP(A) As Variant()
Dim Dr()
Dr = LinMthDr(A)
If Sz(Dr) = 0 Then Exit Function
Dr(3) = AyAddCommaSpcSfxExptLas(AyTrim(SplitComma(Dr(3))))
LinMthDrWP = Dr
End Function

Function LinMthDr(A) As Variant()
Dim L$, Mdy$, Ty$, Nm$, Prm$, Ret$, TopRmk$, LinRmk$
L = A
Mdy = ShfShtMdy(L)
Ty = ShfMthTy(L): If Ty = "" Then Exit Function
Ty = MthShtTy(Ty)
Nm = ShfNm(L)
Ret = ShfMthSfx(L)
Prm = ShfBktStr(L)
If ShfX(L, "As") = "As" Then
    If Ret <> "" Then Stop
    Ret = ShfTerm(L)
End If
If ShfX(L, "'") = "'" Then
    LinRmk = L
End If
LinMthDr = Array(Mdy, Ty, Nm, Prm, Ret, LinRmk)
End Function

Function ShfRmk(A) As String()
Dim L$
L = LTrim(A)
If FstChr(L) = "'" Then
    ShfRmk = ApSy(Mid(L, 2), "")
Else
    ShfRmk = ApSy("", A)
End If
End Function

Function PjMthDryWh(A As VBProject, B As WhMth, C As MthWhOpt) As Variant()
Dim M
For Each M In AyNz(PjMdAy(A))
    PushAy PjMthDryWh, MdMthDryWh(CvMd(M), B, C)
Next
End Function

Function PjMthDry(A As VBProject) As Variant()
Dim M
For Each M In AyNz(PjMdAy(A))
    PushAy PjMthDry, MdMthDry(CvMd(M))
Next
End Function

Function PjMthKyWs(A As VBProject) As Worksheet
Set PjMthKyWs = WsVis(SqWs(PjMthKySq(A)))
End Function

Function PjMthWs(A As CodeModule) As Worksheet
Set PjMthWs = WsVis(SqWs(PjMthSq(A)))
End Function

Function CvVbe(A) As Vbe
Set CvVbe = A
End Function

Sub AcsCls()
On Error Resume Next
Acs.CloseCurrentDatabase
End Sub

Function AcsOpnFb(A$) As Access.Application
AcsCls
Acs.OpenCurrentDatabase A
Set AcsOpnFb = Acs
End Function

Function ZZVbeAy() As Vbe()
PushObj ZZVbeAy, CurVbe
Const Fb$ = "C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipRate\StockShipRate\StockShipRate (ver 1.0).accdb"
PushObj ZZVbeAy, AcsOpnFb(Fb).Vbe
End Function

Sub Z_VbeAyMthWs()
WsVis VbeAyMthWs(ZZVbeAy)
End Sub




Sub PushDrs(O As Drs, A As Drs)
If Not IsEq(O.Fny, A.Fny) Then Stop
Set O = Drs(O.Fny, CvAy(AyAddAp(O.Dry, A.Dry)))
End Sub







Function CurVbeMthWb() As Workbook
Set CurVbeMthWb = VbeMthWb(CurVbe)
End Function



Function LoWs(A As ListObject) As Worksheet
Set LoWs = A.Parent
End Function

Function CurVbeMthWs() As Worksheet
Set CurVbeMthWs = VbeMthWs(CurVbe)
End Function


Private Sub ZZ_SrcMthDry()
Dim A(): A = SrcMthDry(CurSrc)
Stop
End Sub

Private Sub ZZ_VbeMthWb()
WbVis VbeMthWb(CurVbe)
End Sub

Private Sub ZZ_VbeMthWs()
WsVis VbeMthWs(CurVbe)
End Sub

Function LasFilSeg$(A$)
LasFilSeg = AyLasEle(Split(A, "\"))
End Function
Function CutExt$(A$)

End Function


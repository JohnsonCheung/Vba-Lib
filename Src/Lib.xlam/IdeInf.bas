Attribute VB_Name = "IdeInf"
Option Explicit
Enum eSrcTy
   eDtaTy
   eMth
End Enum
Type SrcItm
   SrcTy As eSrcTy
   Nm As String
   Ly() As String
End Type
Enum eLisMdSrt
   elmsLines
   elmsMd
   elmsNMth
End Enum
Type SrcItmCnt
    N As Integer
    NPub As Integer
    NPrv As Integer
End Type
Function FnyOfMthDrs(Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As String()
Dim O$(): O = SplitSpc("MdNm Lno Mdy Ty MthNm")
If WithBdyLy Then Push O, "BdyLy"
If WithBdyLines Then Push O, "BdyLines"
FnyOfMthDrs = O
End Function
Function WinOfImm() As VBIDE.Window
Set WinOfImm = WinTy_Win(vbext_wt_Immediate)
End Function
Function CvWin(V) As VBIDE.Window
Set CvWin = V
End Function
Sub WinAp_Keep(ParamArray Ap())
Dim Av(): Av = Ap
WinAy_Keep Av
End Sub
Function ObjPointer&(V)
ObjPointer = ObjPtr(V)
End Function
Sub WinAy_Keep(A)
Dim W As VBIDE.Window
Dim P&(): P = AyMap_Lng(A, "ObjPointer")
For Each W In CurVbe.Windows
    If Not AyHas(P, ObjPtr(W)) Then
        W.Close
    End If
Next
Dim I
For Each I In A
    CvWin(I).Visible = True
Next
End Sub
Sub WinOfImm_Cls()
DoEvents
With WinOfImm
    .Visible = False
End With
DoEvents
WinOfBrwObj.SetFocus
'Interaction.SendKeys "^{F4}", True
'DoEvents
End Sub
Function VisWinCnt&()
VisWinCnt = ItrCntByBoolPrp(CurVbe.Windows, "Visible")
End Function
Function WinOfBrwObj() As VBIDE.Window
Set WinOfBrwObj = WinTy_Win(vbext_wt_Browser)
End Function
Sub ClrImmWin()
With WinOfImm
    .SetFocus
    .Visible = True
End With
Interaction.SendKeys "^{HOME}^+{END} ", True
End Sub
Sub CmdBarOfMnu__Tst()
Debug.Print CmdBarOfMnu.Name
End Sub
Function CmdBarOfMnu() As CommandBar
Set CmdBarOfMnu = CurVbe.CommandBars("Menu Bar")
End Function
Function CmdBarCap_CmdPop(A As CommandBar, Cap$) As CommandBarPopup
Set CmdBarCap_CmdPop = ItrItmByPrp(A.Controls, "Caption", Cap)
End Function
Sub CmdPopOfWin__Tst()
Debug.Print CmdPopOfWin.Caption
End Sub

Function CmdPopOfWin() As CommandBarPopup
Set CmdPopOfWin = CmdBarCap_CmdPop(CmdBarOfMnu, "&Window")
End Function
Function CmdBtnOfTileV() As CommandBarButton
Set CmdBtnOfTileV = CmdPopCap_CmdBtn(CmdPopOfWin, "Tile &Vertically")
End Function
Function CmdBtnOfTileH() As CommandBarButton
Set CmdBtnOfTileH = CmdPopCap_CmdBtn(CmdPopOfWin, "Tile &Horizontally")
End Function
Sub TileV()
CmdBtnOfTileV.Execute
End Sub
Function CmdPopCap_CmdBtn(A As CommandBarPopup, Cap$) As CommandBarButton
Set CmdPopCap_CmdBtn = ItrItmByPrp(A.Controls, "Caption", Cap)
End Function
Sub TileH()
CmdBtnOfTileH.Execute
End Sub
Function WinTy_Win(A As vbext_WindowType) As VBIDE.Window
Set WinTy_Win = ItrItmByPrp(CurVbe.Windows, "Type", A)
End Function
Function NewSrcItmCnt(N%, NPub%, NPrv%) As SrcItmCnt
With NewSrcItmCnt
    .N = N
    .NPrv = NPrv
    .NPub = NPub
End With
End Function

Function MdMdLisDt(A As CodeModule, Optional MdNmPatn$ = ".", Optional Srt As eLisMdSrt) As Dt
Dim Dt As Dt: ' Dt = MdMdInfDt(A, MdNmPatn)
Select Case Srt
Case elmsLines: Dt = DtSrt(Dt, "Lines", True)
Case elmsMd: Dt = DtSrt(Dt, "Md")
Case elmsNMth: Dt = DtSrt(Dt, "NMth", True)
Case Else: Stop
End Select
MdMdLisDt = Dt
End Function

Sub MdMdLisDtBrw(A As CodeModule, Optional MdNmPatn$ = ".", Optional Srt As eLisMdSrt)
DtBrw MdMdLisDt(A, MdNmPatn, Srt)
End Sub

Sub MdMdLisDtDmp(A As CodeModule, Optional MdNmPatn$ = ".", Optional Srt As eLisMdSrt)
DtDmp MdMdLisDt(A, MdNmPatn, Srt)
End Sub

Function PjPjPrpInfDt(A As VBProject) As Dt

End Function

Function PjMthNmDrs(A As VBProject, CmpTyAy() As vbext_ComponentType, Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".", Optional Sep$ = vbTab) As Drs
PjMthNmDrs = NewDrs("Md Mth", DotNyDry(PjMthNy(A, CmpTyAy, MthNmPatn, MdNmPatn, Sep)))
End Function

Function FnyOf_MdLis() As String()
FnyOf_MdLis = SplitSpc("PJ Md-Pfx Md Ty Lines NMth NMth-Pub NMth-Prv NTy NTy-Pub NTy-Prv NEnm NEnm-Pub NEnm-Prv")
End Function

Function MdMdLisDr(A As CodeModule) As Variant()
Dim O(), Md As CodeModule
Push O, PjNm(MdPj(A))
Push O, TakBef(MdNm(A), "_")
Push O, MdNm(Md)
Push O, MdTyNm(Md)
Push O, Md.CountOfLines
With MdMthCnt(A)
   Push O, .N
   Push O, .NPub
   Push O, .NPrv
End With
With MdTyItmCnt(A)
   Push O, .N
   Push O, .NPub
   Push O, .NPrv
End With
With MdEnmItmCnt(A)
   Push O, .N
   Push O, .NPub
   Push O, .NPrv
End With
MdMdLisDr = O
End Function

Function MdMthCnt(A As CodeModule) As SrcItmCnt
Dim B As Drs: B = MdMthDrs(A)
Dim N%, NPub%, NPrv%
N = Sz(B.Dry)
NPub = DrsRowCnt(B, "Mdy", "Public") + DrsRowCnt(B, "Mdy", "")
NPrv = DrsRowCnt(B, "Mdy", "Private")
MdMthCnt = NewSrcItmCnt(N, NPub, NPrv)
End Function
Function MdMthDry(A As CodeModule, Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Variant()
MdMthDry = SrcMthDry(MdSrc(A), MdNm(A), WithBdyLy, WithBdyLines)
End Function

Function MdMthDrs(A As CodeModule, Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Drs
MdMthDrs = SrcMthDrs(MdSrc(A), MdNm(A), WithBdyLy, WithBdyLines)
End Function
Function SrcLin_MthNmPos%(A)
End Function
Function MdMth_RRCC(A As CodeModule, MthNm$) As RRCC
MdMth_RRCC = SrcMth_RRCC(MdSrc(A), MthNm)
End Function
Function MthDrs_SortingKeyAy(A As Drs) As String()
If AyIsEmp(A.Dry) Then Exit Function
Dim Dr, Mdy$, Ty$, MthNm$, O$()
For Each Dr In DrsSel(A, "Mdy Ty MthNm").Dry
    AyAsg Dr, Mdy, Ty, MthNm
    Push O, MthDrs_SortingKeyAy__CrtKey(Mdy, Ty, MthNm)
Next
MthDrs_SortingKeyAy = O
End Function

Private Function MthDrs_SortingKeyAy__CrtKey$(Mdy$, Ty$, MthNm$)
Dim A1 As Byte
    If HasSfx(MthNm, "__Tst") Then
        A1 = 8
    ElseIf MthNm = "Tst" Then
        A1 = 9
    Else
        Select Case Mdy
        Case "Public", "": A1 = 1
        Case "Friend": A1 = 2
        Case "Private": A1 = 3
        Case Else: Stop
        End Select
    End If
Dim A3$
    If Ty <> "Function" And Ty <> "Sub" Then A3 = Ty
MthDrs_SortingKeyAy__CrtKey = FmtQQ("?:?:?", A1, MthNm, A3)
End Function


Function MdTyItmCnt(A As CodeModule) As SrcItmCnt

End Function

Function PjMdLisDry(A As VBProject) As Variant()
Dim I, M As CodeModule, O()
For Each I In PjMdAy(A)
   Set M = I
   Push O, MdMdLisDr(M)
Next
PjMdLisDry = O
End Function

Function PjMdLisDt(A As VBProject, Optional MdNmPatn$ = ".") As Dt
Dim I, Md As CodeModule
Dim O()
For Each I In PjMdAy(A, MdNmPatn)
   Set Md = I
   Push O, MdMdLisDr(Md)
Next
PjMdLisDt = NewDt("Md", FnyOf_MdLis, O)
End Function


Function NewMthSrc(Nm$, Ly$()) As SrcItm
NewMthSrc.Nm = Nm
NewMthSrc.Ly = Ly
NewMthSrc.SrcTy = eMth
End Function

Function NewTySrc(Nm$, Ly$()) As SrcItm
NewTySrc.Nm = Nm
NewTySrc.Ly = Ly
NewTySrc.SrcTy = eDtaTy
End Function

Sub SrcItmAyDmp(A() As SrcItm)
Dim J%
For J = 0 To SrcItmUB(A)
AyDmp SrcItmLy(A(J))
Next
End Sub

Sub SrcItmDmp(A As SrcItm)
AyDmp SrcItmLy(A)
End Sub

Function SrcItmLy(A As SrcItm) As String()
SrcItmLy = A.Ly
End Function

Sub SrcItmPush(O() As SrcItm, M As SrcItm)
Dim N%: N = SrcItmSz(O)
ReDim Preserve O(N)
O(N) = M
End Sub

Sub SrcItmPushAy(O() As SrcItm, M() As SrcItm)
Dim J%
For J = 0 To SrcItmUB(M)
   SrcItmPush O, M(J)
Next
End Sub

Function SrcItmSz%(A() As SrcItm)
On Error Resume Next
SrcItmSz = UBound(A) + 1
End Function

Function SrcItmUB%(A() As SrcItm)
SrcItmUB = SrcItmSz(A) - 1
End Function

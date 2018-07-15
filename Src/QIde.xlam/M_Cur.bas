Attribute VB_Name = "M_Cur"
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

Property Get CurCdWin() As VBIDE.Window
Stop '
'Set CurCdWin = Vbe.ActiveCodePane.Window
End Property

Property Get CurMd() As VBIDE.CodeModule
Set CurMd = CurCdPne.CodeModule
End Property

Property Get CurMdNm$()
CurMdNm = MdNm(CurMd)
End Property

Property Get CurPj() As VBProject
Set CurPj = CurVbe.ActiveVBProject
End Property

Property Get CurVbe() As Vbe
Set CurVbe = Application.Vbe
End Property

Property Get CurWin() As VBIDE.Window
Set CurWin = CurVbe.ActiveWindow
End Property

Sub ClrImmWin()
With WinOfImm
    .SetFocus
    .Visible = True
End With
Interaction.SendKeys "^{HOME}^+{END} ", True
End Sub

Function CmdBarCap_CmdPop(A As CommandBar, Cap$) As CommandBarPopup
Stop '
'Set CmdBarCap_CmdPop = ItrItmByPrp(A.Controls, "Caption", Cap)
End Function

Function CmdBarOfMnu() As CommandBar
Set CmdBarOfMnu = CurVbe.CommandBars("Menu Bar")
End Function

Function CmdBtnOfTileH() As CommandBarButton
Set CmdBtnOfTileH = CmdPopCap_CmdBtn(CmdPopOfWin, "Tile &Horizontally")
End Function

Function CmdBtnOfTileV() As CommandBarButton
Set CmdBtnOfTileV = CmdPopCap_CmdBtn(CmdPopOfWin, "Tile &Vertically")
End Function

Function CmdPopCap_CmdBtn(A As CommandBarPopup, Cap$) As CommandBarButton
Stop '
'Set CmdPopCap_CmdBtn = ItrItmByPrp(A.Controls, "Caption", Cap)
End Function

Function CmdPopOfWin() As CommandBarPopup
Set CmdPopOfWin = CmdBarCap_CmdPop(CmdBarOfMnu, "&Window")
End Function

Function FnyOfMthDrs(Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As String()
Dim O$(): O = SplitSpc("MdNm Lno Mdy Ty MthNm")
If WithBdyLy Then Push O, "BdyLy"
If WithBdyLines Then Push O, "BdyLines"
FnyOfMthDrs = O
End Function

Function FnyOf_MdLis() As String()
FnyOf_MdLis = SplitSpc("PJ Md-Pfx Md Ty Lines NMth NMth-Pub NMth-Prv NTy NTy-Pub NTy-Prv NEnm NEnm-Pub NEnm-Prv")
End Function

Function MdMdLisDr(A As CodeModule) As Variant()
Dim O(), N$
Stop '
'N = MdNm(A)
'Push O, Pjx(MdPj(A)).Nm
'Push O, TakBef(N, "_")
'Push O, N
'Push O, MdTyNm(A)
'Push O, A.CountOfLines
'With MdMthCnt(A)
'   Push O, .N
'   Push O, .NPub
'   Push O, .NPrv
'End With
'With MdTyItmCnt(A)
'   Push O, .N
'   Push O, .NPub
'   Push O, .NPrv
'End With
'With MdEnmItmCnt(A)
'   Push O, .N
'   Push O, .NPub
'   Push O, .NPrv
'End With
'MdMdLisDr = O
End Function

Function MdMdLisDt(A As CodeModule, Optional MdNmPatn$ = ".", Optional Srt As eLisMdSrt) As Dt
Dim Dt As Dt: ' Dt = MdMdInfDt(A, MdNmPatn)
Stop '
Select Case Srt
'Case elmsLines: Dt = DtSrt(Dt, "Lines", True)
'Case elmsMd: Dt = DtSrt(Dt, "Md")
'Case elmsNMth: Dt = DtSrt(Dt, "NMth", True)
Case Else: Stop
End Select
MdMdLisDt = Dt
End Function

Sub MdMdLisDtBrw(A As CodeModule, Optional MdNmPatn$ = ".", Optional Srt As eLisMdSrt)
Stop '
'DtBrw MdMdLisDt(A, MdNmPatn, Srt)
End Sub

Sub MdMdLisDtDmp(A As CodeModule, Optional MdNmPatn$ = ".", Optional Srt As eLisMdSrt)
Stop '
'DtDmp MdMdLisDt(A, MdNmPatn, Srt)
End Sub

Function MdMthCnt(A As CodeModule) As SrcItmCnt
Dim B As Drs: B = MdMthDrs(A)
Dim N%, NPub%, NPrv%
N = Sz(B.Dry)
Stop '
'NPub = DrsRowCnt(B, "Mdy", "Public") + DrsRowCnt(B, "Mdy", "")
'NPrv = DrsRowCnt(B, "Mdy", "Private")
MdMthCnt = NewSrcItmCnt(N, NPub, NPrv)
End Function

Function MdMthDrs(A As CodeModule, Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Drs
Set MdMthDrs = SrcMthDrs(MdSrc(A), MdNm(A), MdTy(A), WithBdyLy, WithBdyLines)
End Function

Function MdMthDry(A As CodeModule, Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Variant()
Stop '
'MdMthDry = SrcMthDry(MdSrc(A), MdNm(A), MdTyStr(A), WithBdyLy, WithBdyLines)
End Function

Function MdMth_RRCC(A As CodeModule, MthNm$) As RRCC
Stop '
'MdMth_RRCC = SrcMth_RRCC(MdSrc(A), MthNm)
End Function

Function MdTyItmCnt(A As CodeModule) As SrcItmCnt

End Function

Function MthDrs_SortingKy(A As Drs) As String()
If AyIsEmp(A.Dry) Then Exit Function
Dim Dr, Mdy$, Ty$, MthNm$, O$()
Stop '
'For Each Dr In DrsSel(A, "Mdy Ty MthNm").Dry
'    AyAsg Dr, Mdy, Ty, MthNm
'    Push O, MthDrs_SortingKy__CrtKey(Mdy, Ty, MthNm)
'Next
MthDrs_SortingKy = O
End Function

Function NewMthSrc(Nm$, Ly$()) As SrcItm
NewMthSrc.Nm = Nm
NewMthSrc.Ly = Ly
NewMthSrc.SrcTy = eMth
End Function

Function NewSrcItmCnt(N%, NPub%, NPrv%) As SrcItmCnt
With NewSrcItmCnt
    .N = N
    .NPrv = NPrv
    .NPub = NPub
End With
End Function

Function NewTySrc(Nm$, Ly$()) As SrcItm
NewTySrc.Nm = Nm
NewTySrc.Ly = Ly
NewTySrc.SrcTy = eDtaTy
End Function

Function ObjPointer&(V)
ObjPointer = ObjPtr(V)
End Function

Function PjMdLisDry(A As VBProject) As Variant()
Dim I, O()
Stop '
'For Each I In Pjx(A).MdAy
'   Push O, MdMdLisDr(CvMd(I))
'Next
PjMdLisDry = O
End Function

Function PjMdLisDt(A As VBProject, Optional MdNmPatn$ = ".") As Dt
Dim I, Md As CodeModule
Dim O()
Stop '
'For Each I In Pjx(A).MdAy(MdNmPatn)
'   Set Md = I
'   Push O, MdMdLisDr(Md)
'Next
'PjMdLisDt = NewDt("Md", FnyOf_MdLis, O)
End Function

Function PjMthNmDrs(A As VBProject, Optional CmpTyAy0, Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".") As Drs
Stop '
'Dim MthNy$(): MthNy = Pjx(A).MthNy(CmpTyAy0, MthNmPatn, MdNmPatn)
'Dim O(): O = DotNy_Dry(MthNy)
'Stop
'PjMthNmDrs = NewDrs("Md Mth", O)
End Function

Function PjPjPrpInfDt(A As VBProject) As Dt

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

Function SrcLin_MthNmPos%(A)
End Function

Sub TileH()
CmdBtnOfTileH.Execute
End Sub

Sub TileV()
CmdBtnOfTileV.Execute
End Sub

Function VisWinCnt&()
Stop '
'VisWinCnt = ItrCntByBoolPrp(CurVbe.Windows, "Visible")
End Function

Sub WinAp_Keep(ParamArray Ap())
Dim Av(): Av = Ap
WinAy_Keep Av
End Sub

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

Function WinOfBrwObj() As VBIDE.Window
Set WinOfBrwObj = WinTy_Win(vbext_wt_Browser)
End Function

Function WinOfImm() As VBIDE.Window
Set WinOfImm = WinTy_Win(vbext_wt_Immediate)
End Function

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

Function WinTy_Win(A As vbext_WindowType) As VBIDE.Window
Stop '
'Set WinTy_Win = ItrItmByPrp(CurVbe.Windows, "Type", A)
End Function

Sub CmdBarOfMnu__Tst()
Debug.Print CmdBarOfMnu.Name
End Sub

Sub CmdPopOfWin__Tst()
Debug.Print CmdPopOfWin.Caption
End Sub

Private Function MthDrs_SortingKy__CrtKey$(Mdy$, Ty$, MthNm$)
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
MthDrs_SortingKy__CrtKey = FmtQQ("?:?:?", A1, MthNm, A3)
End Function

Private Sub ZZ_PjMthNmDrs()
Stop '
'DrsBrw PjMthNmDrs(CurPj)
End Sub

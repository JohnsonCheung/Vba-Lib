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
'Set CurCdWin = VBE.ActiveCodePane.Window
End Property

Property Get CurMd() As VBIDE.CodeModule
Set CurMd = CurCdPne.CodeModule
End Property

Property Get CurMdNm$()
CurMdNm = MdNm(CurMd)
End Property

Function CurPj() As VBProject
Set CurPj = CurVbe.ActiveVBProject
End Function

Property Get CurVbe() As VBE
Set CurVbe = Application.VBE
End Property

Property Get CurWin() As VBIDE.Window
Set CurWin = CurVbe.ActiveWindow
End Property

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

Sub CurMd__Tst()
Ass CurMd.Parent.Name = "Cur_d"
End Sub
Function CurMthBdyLines$()
CurMthBdyLines = MthBdyLines(CurMd, CurMthNm$)
End Function
Function CurMthNm$()
CurMthNm = MdCurMthNm(CurMd)
End Function
Function CurTarMd() As CodeModule
With CurVbe
   If .CodePanes.Count <> 2 Then Exit Function
   Dim M1 As CodeModule: Set M1 = .CodePanes(1).CodeModule
   Dim M2 As CodeModule: Set M2 = .CodePanes(2).CodeModule
   Dim M As CodeModule: Set M = CurMd
   Dim IsM1Tar As Boolean: IsM1Tar = M1 <> M And M2 = M
   Dim IsM2Tar As Boolean: IsM2Tar = M2 <> M And M1 = M
   If Not (IsM1Tar Xor IsM2Tar) Then Stop
   If IsM1Tar Then Set CurTarMd = M1: Exit Function
   If IsM2Tar Then Set CurTarMd = M2: Exit Function
End With
End Function
Sub CurTarMd__Tst()
Debug.Print MdNm(CurTarMd)
End Sub

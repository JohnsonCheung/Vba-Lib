Attribute VB_Name = "IdeMthDic"
Option Explicit
Function PjDic(A As VBProject) As Dictionary
Dim DicAy() As Dictionary
DicAy = AyMapInto(PjMdAy(A), "MdMthDic", DicAy)
Set PjDic = DicAyAdd(DicAy)
End Function

Private Sub ZZ_PjDic()
DicBrw PjDic(CurPj)
End Sub

Private Sub ZZ_MdDic()
DicBrw MdDic(CurMd)
End Sub

Private Sub ZZ_SrcDic()
'Dim A As Dictionary: Set A = SrcDicOfMthNmzzzMthLines(CurSrc)
DicBrw SrcDic(CurSrc)
End Sub

Private Sub Z_PjDic()
Dim A As Dictionary, V, K
Set A = PjDic(CurPj)
Ass IsSyDic(A) '
For Each K In A
    If InStr(K, ".") > 0 Then Stop
    If Sz(A(K)) = 0 Then Stop
Next
End Sub

Function MdDic(A As CodeModule) As Dictionary
Set MdDic = DicAddKeyPfx(SrcDic(MdSrc(A)), MdDNm(A) & ".")
End Function

Private Sub Z_MdDic()
DicBrw MdDic(CurMd)
End Sub


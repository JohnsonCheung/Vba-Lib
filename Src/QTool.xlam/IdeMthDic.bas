Attribute VB_Name = "IdeMthDic"
Option Explicit


Function PjMthDic(A As VBProject) As Dictionary
Dim DicAy() As Dictionary
DicAy = AyMapInto(PjMdAy(A), "MdMthDic", DicAy)
Set PjMthDic = DicAyAdd(DicAy)
End Function

Private Sub ZZ_PjMthDic()
DicBrw PjMthDic(CurPj)
End Sub

Private Sub ZZ_MdMthDic()
DicBrw MdMthDic(CurMd)
End Sub

Private Sub ZZ_SrcMthDic()
'Dim A As Dictionary: Set A = SrcDicOfMthNmzzzMthLines(CurSrc)
DicBrw SrcMthDic(CurSrc)
End Sub

Private Sub Z_PjMthDic()
Dim A As Dictionary, V, K
Set A = PjMthDic(CurPj)
Ass IsSyDic(A) '
For Each K In A
    If InStr(K, ".") > 0 Then Stop
    If Sz(A(K)) = 0 Then Stop
Next
End Sub




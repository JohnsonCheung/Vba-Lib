Attribute VB_Name = "M_MthDrs"
Option Explicit

Function MthDrs_Ky(A As Drs) As String()
Dim Ty$, Mdy$, MthNm$, K$, IxAy%(), Dr, O$()
IxAy = FnyIxAy(A.Fny, "Mdy MthNm Ty")
If AyIsEmp(A.Dry) Then Exit Function
For Each Dr In A.Dry
    'Debug.Print Mdy, MthNm, Ty
    DrIxAy_Asg Dr, IxAy, Mdy, MthNm, Ty
    K = MthNm & ":" & Ty & ":" & Mdy
    Push O, K
Next
MthDrs_Ky = O
End Function
Sub MthDrs_SortingKy__Tst()
'AyDmp MthDrs_SortingKy(SrcMthDrs(MdSrc(Md("Mth_"))))
End Sub
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

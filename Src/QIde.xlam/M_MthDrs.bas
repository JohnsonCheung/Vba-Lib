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

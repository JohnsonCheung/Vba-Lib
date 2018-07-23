Attribute VB_Name = "M_Dic"
Option Explicit

Function DicDr(A As Dictionary, Ky$()) As Variant()
Dim O(), I, J&
ReDim O(UB(Ky))
For Each I In Ky
    If A.Exists(I) Then
        O(J) = A(I)
    End If
    J = J + 1
Next
DicDr = O
End Function

Function DicDrs(A As Dictionary, Optional InclDicValTy As Boolean) As Drs
Dim Fny$()
Fny = SplitSpc("Key Val"): If InclDicValTy Then Push Fny, "ValTy"
Set DicDrs = Drs(Fny, DicDry(A, InclDicValTy))
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
Dim Dry()
Dry = DicDry(A, InclDicValTy)
Dim F$
    If InclDicValTy Then
        F = "Key Val Ty"
    Else
        F = "Key Val"
    End If
Set DicDt = Dt(DtNm, F, Dry)
End Function

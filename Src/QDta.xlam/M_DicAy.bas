Attribute VB_Name = "M_DicAy"
Option Explicit

Function DicAy_Drs(A() As Dictionary, Optional Fny0) As Drs
Const CSub$ = "DicAy_Drs"
Dim UDic%
   UDic = UB(A)
Dim Fny$()
    Fny = DftNy(Fny0)
    If AyIsEmp(Fny) Then
        Dim J%
        Push Fny, "Key"
        For J = 0 To UDic
            Push Fny, "V" & J
        Next
    Else
        Fny = Fny0
    End If
If UB(Fny) <> UDic + 1 Then Er CSub, "Given {Fny0} has {Sz} <> {DicAy-Sz}", Fny, Sz(Fny), Sz(A)
Dim Ky$()
   Ky = DicAy_Ky(A)
Dim O()
    Dim I
   ReDim O(UDic)
   J = 0
   For Each I In A
       O(J) = DicDr(CvDic(I), Ky)
       J = J + 1
   Next
Set DicAy_Drs = Drs(Fny, O)
End Function

Function DicAy_Ky(A() As Dictionary) As String()
Dim O$(), I
For Each I In A
    PushNoDupAy O, CvDic(I).Keys
Next
DicAy_Ky = O
End Function

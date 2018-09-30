Attribute VB_Name = "VbDic"
Option Explicit
Function DicAB_SamDic(A As Dictionary, B As Dictionary) As Dictionary
Dim O As New Dictionary
If A.Count = 0 Or B.Count = 0 Then GoTo X
Dim K
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) = B(K) Then
            O.Add K, A(K)
        End If
    End If
Next
X: Set DicAB_SamDic = O
End Function
Function DicAB_SamKeyDifVal_DicPair(A As Dictionary, B As Dictionary) As Variant()
Dim K, A1 As New Dictionary, B1 As New Dictionary
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) <> B(K) Then
            A1.Add K, A(K)
            B1.Add K, B(K)
        End If
    End If
Next
DicAB_SamKeyDifVal_DicPair = Array(A1, B1)
End Function


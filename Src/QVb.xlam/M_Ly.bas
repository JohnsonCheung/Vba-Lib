Attribute VB_Name = "M_Ly"
Option Explicit

Function LyDic(A$(), Optional JnSep$ = vbCrLf) As Dictionary
Const CSub$ = "LyDic"
Dim O As New Dictionary
   If AyIsEmp(A) Then Set LyDic = O: Exit Function
   Dim I
   For Each I In A
       If Trim(I) = "" Then GoTo Nxt
       If FstChr(I) = "#" Then GoTo Nxt
       With Brk1(I, " ")
           If O.Exists(.S1) Then
               O(.S1) = O(.S1) & JnSep & .S2
           Else
               O.Add .S1, .S2
           End If
       End With
Nxt:
   Next
Set LyDic = O
End Function

Function LyEndTrim(A$()) As String()
If AyIsEmp(A) Then Exit Function
If AyLasEle(A) <> "" Then LyEndTrim = A: Exit Function
Dim J%
For J = UB(A) To 0 Step -1
    If Not Trim(A(J)) <> "" Then
        Dim O$()
        O = A
        ReDim Preserve O(J)
        LyEndTrim = O
        Exit Function
    End If
Next
End Function

Function LyHasMajPfx(A$(), MajPfx$) As Boolean
Dim Cnt%, J%
For J = 0 To UB(A)
    If HasPfx(A(J), MajPfx) Then Cnt = Cnt + 1
Next
LyHasMajPfx = Cnt > (Sz(A) \ 2)
End Function

Function LyRmv2Dash(A$()) As String()
If Sz(A) = 0 Then Exit Function
Dim O$(), I
For Each I In A
    Push O, Rmv2Dash(CStr(I))
Next
LyRmv2Dash = O
End Function

Function LySqH(A$()) As Variant()
LySqH = AySqH(A)
End Function

Function LySqV(A$()) As Variant()
LySqV = AySqV(A)
End Function

Function LyToStr$(A$())
If Sz(A) = 0 Then
    LyToStr = "Ly()"
Else
    LyToStr = FmtQQ("Ly(|?|)", JnCrLf(A, WithIx:=True))
End If
End Function

Function Ly_T1Rst_SyPair(A$()) As SyPair
Dim J&, T1$(), Rst$()
For J = 0 To UB(A)
    With LinT1Rst(A(J))
        Push T1, .T1
        Push Rst, .Rst
    End With
Next
Set Ly_T1Rst_SyPair = SyPair(T1, Rst)
End Function

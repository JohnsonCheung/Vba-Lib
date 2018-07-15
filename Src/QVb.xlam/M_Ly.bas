Attribute VB_Name = "M_Ly"
Option Explicit

Property Get LyDic(A$(), Optional JnSep$ = vbCrLf) As Dictionary
Const CSub$ = "LyDic"
Dim O As New Dictionary
   If AyIsEmp(A) Then Set LyDic = O: Exit Property
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
End Property

Property Get LyEndTrim(A$()) As String()
If AyIsEmp(A) Then Exit Property
If AyLasEle(A) <> "" Then LyEndTrim = A: Exit Property
Dim J%
For J = UB(A) To 0 Step -1
    If Not Trim(A(J)) <> "" Then
        Dim O$()
        O = A
        ReDim Preserve O(J)
        LyEndTrim = O
        Exit Property
    End If
Next
End Property

Property Get LyHasMajPfx(A$(), MajPfx$) As Boolean
Dim Cnt%, J%
For J = 0 To UB(A)
    If HasPfx(A(J), MajPfx) Then Cnt = Cnt + 1
Next
LyHasMajPfx = Cnt > (Sz(A) \ 2)
End Property

Property Get LyRmv2Dash(A$()) As String()
If Sz(A) = 0 Then Exit Property
Dim O$(), I
For Each I In A
    Push O, Rmv2Dash(CStr(I))
Next
LyRmv2Dash = O
End Property

Property Get LySqH(A$()) As Variant()
LySqH = AySqH(A)
End Property

Property Get LySqV(A$()) As Variant()
LySqV = AySqV(A)
End Property

Property Get LyToStr$(A$())
If Sz(A) = 0 Then
    LyToStr = "Ly()"
Else
    LyToStr = FmtQQ("Ly(|?|)", JnCrLf(A, WithIx:=True))
End If
End Property

Property Get Ly_T1Rst_SyPair(A$()) As SyPair
Dim J&, T1$(), Rst$()
For J = 0 To UB(A)
    With LinT1Rst(A(J))
        Push T1, .T1
        Push Rst, .Rst
    End With
Next
Set Ly_T1Rst_SyPair = SyPair(T1, Rst)
End Property

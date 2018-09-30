Attribute VB_Name = "VbStrTak"
Option Explicit
Function TakBefOrAll$(S, Sep, Optional NoTrim As Boolean)
TakBefOrAll = Brk1(S, Sep, NoTrim).S1
End Function
Function TakAftOrAll$(S, Sep, Optional NoTrim As Boolean)
TakAftOrAll = Brk2(S, Sep, NoTrim).S2
End Function
Function TakAftMust$(A, Sep, Optional NoTrim As Boolean)
TakAftMust = Brk(A, Sep, NoTrim).S2
End Function
Function TakAft$(A, Sep, Optional NoTrim As Boolean)
TakAft = Brk1(A, Sep, NoTrim).S2
End Function
Function TakBef$(S, Sep$, Optional NoTrim As Boolean)
TakBef = Brk2(S, Sep, NoTrim).S1
End Function
Function TakBefMust$(S, Sep$, Optional NoTrim As Boolean)
TakBefMust = Brk(S, Sep, NoTrim).S1
End Function
Function TakMdy$(A)
TakMdy = TakPfxAyS(A, MdyAy)
End Function
Function TakNm$(A)
Dim J%
If Not IsLetter(Left(A, 1)) Then Exit Function
For J = 2 To Len(A)
    If Not IsNmChr(Mid(A, J, 1)) Then
        TakNm = Left(A, J - 1)
        Exit Function
    End If
Next
TakNm = A
End Function
Function TakPfx$(Lin, Pfx$) ' Return [Pfx] if [Lin] has such pfx else return ""
If HasPfx(Lin, Pfx) Then TakPfx = Pfx
End Function
Function TakPfxS$(Lin, Pfx$) ' Return [Pfx] if [Lin] has such pfx+" " else return ""
If HasPfx(Lin, Pfx) Then If Mid(Lin, Len(Pfx) + 1, 1) = " " Then TakPfxS = Pfx
End Function
Function TakPfxAyS$(Lin, PfxAy$()) ' Return Fst ele-P of [PfxAy] if [Lin] has pfx ele-P and a space
TakPfxAyS = PfxAyFstS$(PfxAy, Lin)
End Function
Function TakPfxAy$(Lin, PfxAy$()) ' Return Fst ele-P of [PfxAy] if [Lin] has pfx ele-P
TakPfxAy = PfxAyFst$(PfxAy, Lin)
End Function
Function TakBet$(A, S1$, S2$)
Dim P%, L%, P1%, P2%
P1 = InStr(A, S1): If P1 = 0 Then Exit Function
P = P1 + Len(S1)
P2 = InStr(P, A, S2): If P2 = 0 Then Exit Function
L = P2 - P1 - 1
TakBet = Mid(A, P, L)
End Function
Function TakT1$(A)
If FstChr(A) <> "[" Then TakT1 = TakBef(A, " "): Exit Function
Dim P%
P = InStr(A, "]")
If P = 0 Then Stop
TakT1 = Mid(A, 2, P - 2)
End Function
Function TakMthTy$(A)
TakMthTy = TakPfxAy(A, MthTyAy)
End Function
Function TakMthKd$(A)
TakMthKd = TakPfxAyS(A, MthKdAy)
End Function
Function TakMthShtTy$(A)
Dim B$
B = TakPfxAy(A, MthTyAy): If B = "" Then Exit Function
TakMthShtTy = MthShtTy(B)
End Function
Function TakAftBkt$(A)
Dim P%(): P = BktPos(A)
If Sz(P) = 0 Then Exit Function
TakAftBkt = Mid(A, P(1) + 1)
End Function

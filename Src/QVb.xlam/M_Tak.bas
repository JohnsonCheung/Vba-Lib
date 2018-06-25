Attribute VB_Name = "M_Tak"
Option Explicit

Property Get TakAft$(S, Sep, Optional NoTrim As Boolean)
TakAft = Brk1(S, Sep, NoTrim).S2
End Property

Property Get TakAftBkt$(S, Optional Bkt$ = "()")
Dim P2&
   P2 = BrkBktPos(S, Bkt).ToPos
If P2 = 0 Then Exit Property
TakAftBkt = Mid(S, P2 + 1)
End Property

Property Get TakAftRev$(S, Sep, Optional NoTrim As Boolean)
TakAftRev = Brk1Rev(S, Sep, NoTrim).S2
End Property

Property Get TakBef$(S, Sep, Optional NoTrim As Boolean)
TakBef = Brk2(S, Sep, NoTrim).S1
End Property

Property Get TakBefBkt$(S, Optional Bkt$ = "()")
Dim P1&
   P1 = BrkBktPos(S, Bkt).FmPos
If P1 = 0 Then Exit Property
TakBefBkt = Left(S, P1 - 1)
End Property

Property Get TakBefRev$(S, Sep, Optional NoTrim As Boolean)
TakBefRev = BrkRev(S, Sep, NoTrim).S1
End Property

Property Get TakBet$(S, S1, S2, Optional NoTrim As Boolean, Optional InclMarker As Boolean)
With Brk1(S, S1, NoTrim)
   If .S2 = "" Then Exit Property
   Dim O$: O = Brk1(.S2, S2, NoTrim).S1
   If InclMarker Then O = S1 & O & S2
   TakBet = O
End With
End Property

Property Get TakBetBkt$(S, Optional Bkt$ = "()")
Dim P1&, P2&
   With BrkBktPos(S, Bkt)
       P1 = .FmPos
       P2 = .ToPos
   End With
If P1 = 0 Then Exit Property
TakBetBkt = Mid(S, P1 + 1, P2 - P1 - 1)
End Property

Sub ZZ__Tst()
ZZ_TakBet
ZZ_TakBetBkt
End Sub

Private Sub ZZ_TakBet()
Const S1$ = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??"
Const S2$ = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??;AA=XX"
Ass TakBet(S1, "DATABASE=", ";") = "??"
Ass TakBet(S2, "DATABASE=", ";") = "??"
End Sub

Private Sub ZZ_TakBetBkt()
Dim Act$
   Dim S$
   S = "sdklfjdsf(1234()567)aaa("
   Act = TakBetBkt(S)
   Ass Act = "1234()567"
End Sub

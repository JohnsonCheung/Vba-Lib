Attribute VB_Name = "VbStrTak"
Option Explicit

Function TakAft$(S, Sep, Optional NoTrim As Boolean)
TakAft = Brk1(S, Sep, NoTrim).S2
End Function

Function TakAftBkt$(S, Optional Bkt$ = "()")
Dim P2&
   P2 = BktPos(S, Bkt).ToPos
If P2 = 0 Then Exit Function
TakAftBkt = Mid(S, P2 + 1)
End Function

Function TakAftRev$(S, Sep, Optional NoTrim As Boolean)
TakAftRev = Brk1Rev(S, Sep, NoTrim).S2
End Function

Function TakBef$(S, Sep, Optional NoTrim As Boolean)
TakBef = Brk2(S, Sep, NoTrim).S1
End Function

Function TakBefBkt$(S, Optional Bkt$ = "()")
Dim P1&
   P1 = BktPos(S, Bkt).FmPos
If P1 = 0 Then Exit Function
TakBefBkt = Left(S, P1 - 1)
End Function

Function TakBefRev$(S, Sep, Optional NoTrim As Boolean)
TakBefRev = BrkRev(S, Sep, NoTrim).S1
End Function

Function TakBet$(S, S1, S2, Optional NoTrim As Boolean, Optional InclMarker As Boolean)
With Brk1(S, S1, NoTrim)
   If .S2 = "" Then Exit Function
   Dim O$: O = Brk1(.S2, S2, NoTrim).S1
   If InclMarker Then O = S1 & O & S2
   TakBet = O
End With
End Function

Function TakBetBkt$(S, Optional Bkt$ = "()")
Dim P1&, P2&
   With BktPos(S, Bkt)
       P1 = .FmPos
       P2 = .ToPos
   End With
If P1 = 0 Then Exit Function
TakBetBkt = Mid(S, P1 + 1, P2 - P1 - 1)
End Function

Private Sub TakBetBkt__Tst()
Dim Act$
   Dim S$
   S = "sdklfjdsf(1234()567)aaa("
   Act = TakBetBkt(S)
   Ass Act = "1234()567"
End Sub

Private Sub TakBet__Tst()
Const S1$ = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??"
Const S2$ = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??;AA=XX"
Ass TakBet(S1, "DATABASE=", ";") = "??"
Ass TakBet(S2, "DATABASE=", ";") = "??"
End Sub

